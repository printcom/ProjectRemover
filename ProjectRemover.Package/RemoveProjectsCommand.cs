using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ProjectRemover.Package.Classes;
using ProjectRemover.Package.Windows;
using Task = System.Threading.Tasks.Task;

namespace ProjectRemover.Package
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class RemoveProjectsCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int COMMAND_ID = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("1aabcdda-05fa-4f77-b884-a674842ddbe3");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage _package;


        private const string GUID_MATCH = "{([0-9A-Fa-f]{8}[-][0-9A-Fa-f]{4}[-][0-9A-Fa-f]{4}[-][0-9A-Fa-f]{4}[-][0-9A-Fa-f]{12})}";

        /// <summary>
        /// In the .sln file the projects are linked as shown below. The project guid ist the
        /// same as the guid in the project file.
        /// Project("{9A11103F-16F1-4668-BE54-9A1E7A4F1556}") = "[Name]", "[Path]", "{[Project guid]}"
        /// EndProject
        /// </summary>
        private readonly Regex _referencedProjectsInSolutionRegex = new Regex(
            $"Project.*? = \"(.*?)\", \"(.*?)\", \"{GUID_MATCH}.*?EndProject",
            RegexOptions.Compiled | RegexOptions.Singleline | RegexOptions.IgnoreCase);

        /// <summary>
        /// In the.csproj files the referenced projects are specified in the following form:
        /// <ProjectReference Include ="[Path to csproj file]" >
        /// < Project >{Guid}</Project>
        ///     <Name>Tools</Name>
        /// </ProjectReference>
        ///        
        /// The Guid is the same as the Guid in the.sln file.
        /// </summary>
        private readonly Regex _projectReferenceRegex = new Regex(
            $"<ProjectReference Include=\"(.*?)\">\\s*<Project>{GUID_MATCH}<\\/Project>",
            RegexOptions.Compiled | RegexOptions.Singleline | RegexOptions.IgnoreCase);

        private readonly Regex _numberOfProjectsInSolutionRegex = new Regex("SccNumberOfProjects = ([0-9]+)", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>
        /// Initializes a new instance of the <see cref="RemoveProjectsCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private RemoveProjectsCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            _package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandId = new CommandID(CommandSet, COMMAND_ID);
            var menuItem = new MenuCommand(ExecuteRemoveProjects, menuCommandId);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static RemoveProjectsCommand Instance { get; private set; }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get { return _package; }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in RemoveProjectsCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Instance = new RemoveProjectsCommand(package, commandService);
        }

        #region Event Handler

        private async void ExecuteRemoveProjects(object sender, EventArgs e)
        {
            try
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(_package.DisposalToken);

                WriteToBuildOutputWindow(Strings.info_StartingProjectRemover);

                var dte2 = await ServiceProvider.GetServiceAsync(typeof(DTE)) as DTE2;
                var solution = dte2?.Solution;

                if (solution?.FullName == null)
                {
                    return;
                }

                if (dte2.Solution.SolutionBuild?.BuildState == vsBuildState.vsBuildStateInProgress)
                {
                    WriteToBuildOutputWindow(Strings.warning_RemoveProjectsWhileBuilding);
                    return;
                }

                Dictionary<Guid, (string projectRelativePath, string nestedProjectPath, FileInfo fileInfo)> unusedProjects = GetUnusedProjects(solution.FullName, out Dictionary<Guid, string> solutionFolderNames);

                if (unusedProjects == null ||
                    unusedProjects.Count == 0)
                {
                    WriteToBuildOutputWindow(Strings.info_NoUnusedProjectFound);
                    return;
                }

                RemoveProjectsWindow removeProjectsWindow = new RemoveProjectsWindow();

                foreach (var unusedProject in unusedProjects)
                {
                    removeProjectsWindow.ViewModel.RemovableProjects.Add(new RemovableProject
                    {
                        Id = unusedProject.Key,
                        Name = unusedProject.Value.fileInfo.Name,
                        RelativePath = unusedProject.Value.projectRelativePath,
                        FullPath = unusedProject.Value.fileInfo.FullName,
                        NestedPath = unusedProject.Value.nestedProjectPath,
                        Remove = true
                    });
                }

                removeProjectsWindow.ShowDialog();

                if (removeProjectsWindow.ViewModel.IsCanceled)
                {
                    WriteToBuildOutputWindow(Strings.info_CanceledByUser);
                    return;
                }

                StringBuilder resultText = new StringBuilder();
                int removedProjectsIndex = 0;
                var solutionFileContent = File.ReadAllText(solution.FullName);

                foreach (var removableProject in removeProjectsWindow.ViewModel.RemovableProjects)
                {
                    if (!removableProject.Remove)
                    {
                        continue;
                    }

                    solutionFileContent = RemoveItemFromSolutionFile(solutionFileContent, removableProject.Id, true, removableProject.RelativePath);

                    resultText.Append($"{++removedProjectsIndex}. {removableProject.NestedPath}{removableProject.Name}");
                    resultText.AppendLine();
                }

                int removedSolutionFolders = 0;

                if (removeProjectsWindow.ViewModel.DeleteEmptySolutionFolders)
                {
                    resultText.AppendLine();
                    resultText.Append(Strings.info_RemovedSolutionsFolders);
                    resultText.AppendLine();

                    int solutionFoldersCount = solutionFolderNames.Count;

                    solutionFileContent = DeleteEmptySolutionFolders(solutionFileContent, solutionFolderNames, resultText);
                    removedSolutionFolders = solutionFoldersCount - solutionFolderNames.Count;
                }
                else if (removedProjectsIndex == 0)
                {
                    // No project was removed.
                    return;
                }

                // The file is not locked, so we can override ist.
                File.WriteAllText(solution.FullName, RemoveEmptyLines(solutionFileContent));

                resultText.Insert(0, string.Format(Strings.info_RemovedProjects, removedProjectsIndex, removedSolutionFolders, Environment.NewLine));
                WriteToBuildOutputWindow(resultText.ToString());
            }
            catch (Exception ex)
            {
                WriteToBuildOutputWindow(Strings.error);
                WriteToBuildOutputWindow(ex.Message);
            }
            finally
            {
                WriteToBuildOutputWindow(Strings.info_FinishedProjectRemover);
            }
        }

        #endregion Event Handler

        #region Private Methods

        private void WriteToBuildOutputWindow(string message)
        {
            if (Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(SVsOutputWindow)) is IVsOutputWindow outputWindow)
            {
                Guid generalPaneGuid = VSConstants.BuildOutput;
                outputWindow.GetPane(ref generalPaneGuid, out var generalPane);

                if (generalPane != null)
                {
                    generalPane.OutputString(message);
                    generalPane.OutputString(Environment.NewLine);
                    generalPane.Activate();
                }
            }
        }

        private Dictionary<Guid, (string projectRelativePath, string nestedProjectPath, FileInfo fileInfo)> GetUnusedProjects(string solutionFilePath, out Dictionary<Guid, string> solutionFolderNames)
        {
            solutionFolderNames = new Dictionary<Guid, string>();
            var directoryPath = Path.GetDirectoryName(solutionFilePath);

            if (directoryPath == null ||
                !File.Exists(solutionFilePath))
            {
                return null;
            }

            var fileContent = File.ReadAllText(solutionFilePath);
            var referencedProjectsMatches = _referencedProjectsInSolutionRegex.Matches(fileContent);

            var projectsInSolution = new Dictionary<Guid, (string relativeProjectPath, FileInfo fileInfo)>();

            foreach (Match match in referencedProjectsMatches)
            {
                var relativeFilePath = match.Groups[2].Value;
                var guidValue = match.Groups[3].Value;

                Guid guid = Guid.Parse(guidValue);

                if (!relativeFilePath.EndsWith(".csproj"))
                {
                    // The first match is the name of the solution folder or project
                    solutionFolderNames[guid] = match.Groups[1].Value;

                    // No project. Has to be a solution folder which we don't wanna check.
                    continue;
                }

                var projectPath = Path.Combine(directoryPath, relativeFilePath);

                projectsInSolution.Add(guid, (relativeFilePath, new FileInfo(projectPath)));
            }

            // We add all projects to this collection and remove the ones which are needed. 
            // After that we have only the projects left, which can be removed.
            var unusedProjects = new Dictionary<Guid, (string projectRelativePath, string nestedProjectPath, FileInfo fileInfo)>();

            foreach (var keyValuePair in projectsInSolution)
            {
                var nestedProjectPath = GetNestedProjectPath(fileContent, keyValuePair.Key, solutionFolderNames);
                nestedProjectPath += "/";

                unusedProjects.Add(keyValuePair.Key, (keyValuePair.Value.relativeProjectPath, nestedProjectPath, keyValuePair.Value.fileInfo));
            }

            // Now we need to define the root projects from which we start to check which projects they reference.
            // Root projects are all projects which are saved in the folder of the .sln file or a subfolder from it.
            // We assume that projects, which are saved in a different location, are external projects which were
            // added to the solution.
            var rootProjects = new Dictionary<Guid, (string projectRelativePath, string nestedProjectPath, FileInfo fileInfo)>();

            foreach (var projectKeyValuePair in unusedProjects)
            {
                if (projectKeyValuePair.Value.projectRelativePath.StartsWith(@"..\"))
                {
                    continue;
                }

                rootProjects.Add(projectKeyValuePair.Key, projectKeyValuePair.Value);
            }

            foreach (var rootProjectKeyValuePair in rootProjects)
            {
                // Root projects aren't removed.
                unusedProjects.Remove(rootProjectKeyValuePair.Key);

                bool wasSuccessful = RecursiveReferencedProjectCheck(rootProjectKeyValuePair.Value.fileInfo.FullName, unusedProjects);

                if (!wasSuccessful)
                {
                    return null;
                }
            }

            return unusedProjects;
        }

        private string GetNestedProjectPath(string solutionFileContent, Guid guid, Dictionary<Guid, string> solutionFolderNames)
        {
            string nestedPath = string.Empty;

            Regex nestedProjectRegex = new Regex("{" + guid + "} = " + GUID_MATCH, RegexOptions.IgnoreCase | RegexOptions.Singleline);
            var match = nestedProjectRegex.Match(solutionFileContent);

            if (match.Success &&
                match.Groups.Count > 1 &&
                Guid.TryParse(match.Groups[1].Value, out Guid solutionFolderGuid))
            {
                if (solutionFolderNames.ContainsKey(solutionFolderGuid))
                {
                    nestedPath = solutionFolderNames[solutionFolderGuid];
                }

                nestedPath = GetNestedProjectPath(solutionFileContent, solutionFolderGuid, solutionFolderNames) + $"/{nestedPath}";
            }

            return nestedPath;
        }

        private bool RecursiveReferencedProjectCheck(
            string projectFilePath,
            Dictionary<Guid, (string uniqueProjectName, string nestedProjectPath, FileInfo fileInfo)> unusedProjects)
        {
            if (!File.Exists(projectFilePath))
            {
                // If the project file doesn't exists, we cannot continue. Maybe this project has a reference
                // to a other project which no other project has. Before we remove a wrong project we better
                // don't remove any project.
                return false;
            }

            var projectFileContent = File.ReadAllText(projectFilePath);
            var referencedProjectMatches = _projectReferenceRegex.Matches(projectFileContent);

            foreach (Match referencedProjectMatch in referencedProjectMatches)
            {
                // It turned out that the Project Guid does not match the guid in the solution in some projects (for the same project!).
                // Therefore the complete path to the .csproj file is compared. This must be unique for each project.
                var referencedProjectRelativePath = referencedProjectMatch.Groups[1].Value;
                var currentProjectFileInfo = new FileInfo(Path.Combine(Path.GetDirectoryName(projectFilePath), referencedProjectRelativePath));

                var searchedProject = unusedProjects.Where(p => p.Value.fileInfo.FullName == currentProjectFileInfo.FullName).ToList();

                if (searchedProject.Count > 0)
                {
                    // Now check recursive if this project has any references.
                    bool wasSuccessful = RecursiveReferencedProjectCheck(searchedProject[0].Value.fileInfo.FullName, unusedProjects);

                    if (!wasSuccessful)
                    {
                        return false;
                    }

                    unusedProjects.Remove(searchedProject[0].Key);
                }
            }

            return true;
        }

        private string RemoveItemFromSolutionFile(string solutionFileContent, Guid itemGuid, bool isProject, string projectRelativePath)
        {
            // Remove the project reference.
            solutionFileContent = Regex.Replace(solutionFileContent, $"Project.*{itemGuid}" + "}\"" + @"[\n\r\s]+EndProject", string.Empty, RegexOptions.IgnoreCase);

            // Remove the build configuration for this project.
            solutionFileContent = Regex.Replace(solutionFileContent, "{" + itemGuid + "}.*\\|Any CPU", string.Empty, RegexOptions.IgnoreCase);

            // Remove this project from the solution folder.
            solutionFileContent = Regex.Replace(solutionFileContent, "{" + itemGuid + "} = " + $"{GUID_MATCH}", string.Empty, RegexOptions.IgnoreCase);

            // If this project is linked with tfs, we have to delete some extra entries.
            if (isProject && Regex.IsMatch(solutionFileContent, @"GlobalSection\(TeamFoundationVersionControl\)", RegexOptions.IgnoreCase | RegexOptions.Compiled))
            {
                var numberOfProjects = int.Parse(_numberOfProjectsInSolutionRegex.Match(solutionFileContent).Groups[1].Value);

                // Every project has it's own id which is written after the name.
                // The SccProjectUniqueName is the same as the relative path but with two "\" instead of one.
                // The one will be replaced with four because of the regex.
                var currentProjectNumberRegex = new Regex($"SccProjectUniqueName([0-9]+) = {projectRelativePath.Replace(@"\", @"\\\\")}");

                if (!int.TryParse(currentProjectNumberRegex.Match(solutionFileContent).Groups[1].Value, out int currentProjectNumber))
                {
                    WriteToBuildOutputWindow(Strings.warning);
                    WriteToBuildOutputWindow(Strings.warning_CouldNotDeleteProjectFromTeamFoundationVersionControlSection);

                    return solutionFileContent;
                }

                // Remove all entries for this project.
                solutionFileContent = Regex.Replace(solutionFileContent, $"SccProjectUniqueName{currentProjectNumber} = .*", string.Empty, RegexOptions.IgnoreCase);
                solutionFileContent = Regex.Replace(solutionFileContent, $"SccProjectTopLevelParentUniqueName{currentProjectNumber} = .*", string.Empty, RegexOptions.IgnoreCase);
                solutionFileContent = Regex.Replace(solutionFileContent, $"SccProjectName{currentProjectNumber} = .*", string.Empty, RegexOptions.IgnoreCase);
                solutionFileContent = Regex.Replace(solutionFileContent, $"SccLocalPath{currentProjectNumber} = .*", string.Empty, RegexOptions.IgnoreCase);

                // We need to change the number of all following projects. Otherwise we would get a warning when the solution is loaded,
                // that a project is missing.
                for (int i = currentProjectNumber + 1; i < numberOfProjects; i++)
                {
                    solutionFileContent = Regex.Replace(solutionFileContent, $"SccProjectUniqueName{i}", $"SccProjectUniqueName{i - 1}", RegexOptions.IgnoreCase);
                    solutionFileContent = Regex.Replace(solutionFileContent, $"SccProjectTopLevelParentUniqueName{i}", $"SccProjectTopLevelParentUniqueName{i - 1}", RegexOptions.IgnoreCase);
                    solutionFileContent = Regex.Replace(solutionFileContent, $"SccProjectName{i}", $"SccProjectName{i - 1}", RegexOptions.IgnoreCase);
                    solutionFileContent = Regex.Replace(solutionFileContent, $"SccLocalPath{i}", $"SccLocalPath{i - 1}", RegexOptions.IgnoreCase);
                }

                // We removed one project so we need to decrease the the SccNumberOfProjects.
                solutionFileContent = Regex.Replace(solutionFileContent, $"SccNumberOfProjects = {numberOfProjects}", $"SccNumberOfProjects = {numberOfProjects - 1}", RegexOptions.IgnoreCase);
            }

            return solutionFileContent;
        }

        private string DeleteEmptySolutionFolders(string solutionFileContent, Dictionary<Guid, string> solutionFolderNames, StringBuilder outputInfoBuilder)
        {
            bool removedSolutionFolder = false;

            foreach (var solutionFolderValuePair in solutionFolderNames.ToList())
            {
                // Check if there is still a Item in this solution folder
                Regex nestedItemRegex = new Regex($"{GUID_MATCH} = " + "{" + solutionFolderValuePair.Key + "}", RegexOptions.Singleline | RegexOptions.IgnoreCase);

                if (!nestedItemRegex.IsMatch(solutionFileContent))
                {
                    solutionFileContent = RemoveItemFromSolutionFile(solutionFileContent, solutionFolderValuePair.Key, false, null);
                    removedSolutionFolder = true;
                    solutionFolderNames.Remove(solutionFolderValuePair.Key);
                    outputInfoBuilder.Append(solutionFolderValuePair.Value);
                    outputInfoBuilder.AppendLine();
                }
            }

            if (removedSolutionFolder)
            {
                // If a solution folder was removed, we have to check all other folders again.
                // It's possible that the removed folder was the last item in an other folder.
                // So this folder is now empty and can be removed too.
                solutionFileContent = DeleteEmptySolutionFolders(solutionFileContent, solutionFolderNames, outputInfoBuilder);
            }

            return solutionFileContent;
        }

        /// <summary>
        /// Removes the empty lines in the content of the file. 
        /// </summary>
        private string RemoveEmptyLines(string content)
        {
            // Remove the empty lines.
            var splitContent = content.Split(Environment.NewLine.ToCharArray()).ToList();
            StringBuilder splitContentBuilder = new StringBuilder();

            foreach (var line in splitContent)
            {
                // Sometimes lines only contain "\t\t" and nothing else.
                // These lines are empty too.
                if (string.IsNullOrEmpty(line) ||
                    line == "\t\t")
                {
                    continue;
                }

                splitContentBuilder.Append(line);

                // We removed the line breaks. Add a new one after each line.
                splitContentBuilder.AppendLine();
            }

            return splitContentBuilder.ToString();
        }

        #endregion Private Methods
    }
}
