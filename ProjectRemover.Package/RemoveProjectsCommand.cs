using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
        /// 
        /// </summary>
        readonly Regex _referencedProjectsInSolutionRegex = new Regex(
            $"Project.*? = \".*?\", \"(.*?)\", \"{GUID_MATCH}.*?EndProject",
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
        readonly Regex _projectReferenceRegex = new Regex(
            $"<ProjectReference Include=\"(.*?)\">\\s*<Project>{GUID_MATCH}<\\/Project>",
            RegexOptions.Compiled | RegexOptions.Singleline | RegexOptions.IgnoreCase);

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
        public static RemoveProjectsCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return _package;
            }
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

                Dictionary<Guid, (string uniqueProjectName, string absoluteProjectPath)> unusedProjects = CheckSolutionFile(solution.FullName);

                if (unusedProjects == null ||
                    unusedProjects.Count == 0)
                {
                    // 
                    return;
                }

                var projectEnumerator = solution.Projects.GetEnumerator();
                var allProjectsInSolution = new List<Project>();

                while (projectEnumerator.MoveNext())
                {
                    if (projectEnumerator.Current is Project project)
                    {
                        if (project.Kind == ProjectKinds.vsProjectKindSolutionFolder)
                        {
                            allProjectsInSolution.AddRange(GetProjectsInSolutionFolderRecursive(project));
                        }
                        else
                        {
                            allProjectsInSolution.Add(project);
                        }
                    }
                }

                StringBuilder resultText = new StringBuilder();
                int removedProjectsIndex = 0;

                foreach (var project in allProjectsInSolution.ToList())
                {
                    foreach (var unusedProject in unusedProjects)
                    {
                        if (project.UniqueName == unusedProject.Value.uniqueProjectName)
                        {
                            resultText.Append(++removedProjectsIndex + ". " + project.Name);
                            resultText.AppendLine();

                            solution.Remove(project);

                            // Don't need to check this project again -> remove.
                            unusedProjects.Remove(unusedProject.Key);

                            break;
                        }
                    }
                }

                resultText.Insert(0, string.Format(Strings.info_RemovedProjects, removedProjectsIndex, Environment.NewLine));
                WriteToBuildOutputWindow(resultText.ToString());

            }
            catch (Exception ex)
            {
                WriteToBuildOutputWindow(Strings.error_RemovingProjects);
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

        private IEnumerable<Project> GetProjectsInSolutionFolderRecursive(Project solutionFolder)
        {
            Debug.Assert(solutionFolder.Kind == ProjectKinds.vsProjectKindSolutionFolder, "Incorrect Project type.");

            List<Project> projects = new List<Project>();

            for (var i = 1; i <= solutionFolder.ProjectItems.Count; i++)
            {
                var subProject = solutionFolder.ProjectItems.Item(i).SubProject;

                if (subProject == null)
                {
                    continue;
                }

                // A solution folder has a unique Kind. Recursive check all projects inside.
                if (subProject.Kind == ProjectKinds.vsProjectKindSolutionFolder)
                {
                    projects.AddRange(GetProjectsInSolutionFolderRecursive(subProject));
                }
                else
                {
                    projects.Add(subProject);
                }
            }

            return projects;
        }

        private Dictionary<Guid, (string uniqueProjectName, string absoluteProjectPath)> CheckSolutionFile(string solutionFilePath)
        {
            var directoryPath = Path.GetDirectoryName(solutionFilePath);

            if (directoryPath == null ||
                !File.Exists(solutionFilePath))
            {
                return null;
            }

            var fileContent = File.ReadAllText(solutionFilePath);
            var referencedProjectsMatches = _referencedProjectsInSolutionRegex.Matches(fileContent);

            var projectsInSolution = new Dictionary<Guid, (string relativeProjectPath, string absoluteProjectPath)>();

            foreach (Match match in referencedProjectsMatches)
            {
                var relativeFilePath = match.Groups[1].Value;
                var guidValue = match.Groups[2].Value;

                Guid guid = Guid.Parse(guidValue);

                if (!relativeFilePath.EndsWith(".csproj"))
                {
                    // No project. Has to be a solution folder which we don't wanna check.
                    continue;
                }

                var projectPath = Path.Combine(directoryPath, relativeFilePath);

                projectsInSolution.Add(guid, (relativeFilePath, projectPath));
            }

            // We add all projects to this collection and remove the ones which are needed. 
            // After that we have only the projects left, which can be removed.
            var unusedProjects = new Dictionary<Guid, (string uniqueProjectName, string absoluteProjectPath)>();

            foreach (var keyValuePair in projectsInSolution)
            {
                unusedProjects.Add(keyValuePair.Key, keyValuePair.Value);
            }

            // Now we need to define the root projects from which we start to check which projects they reference.
            // Root projects are all projects which are saved in the folder of the .sln file or a subfolder from it.
            // We assume that projects, which are saved in a different location, are external projects which were
            // added to the solution.
            var rootProjects = new Dictionary<Guid, (string uniqueProjectName, string absoluteProjectPath)>();

            foreach (var projectKeyValuePair in unusedProjects)
            {
                if (projectKeyValuePair.Value.uniqueProjectName.StartsWith(@"..\"))
                {
                    continue;
                }

                rootProjects.Add(projectKeyValuePair.Key, projectKeyValuePair.Value);
            }

            foreach (var rootProjectKeyValuePair in rootProjects)
            {
                // Root projects aren't removed.
                unusedProjects.Remove(rootProjectKeyValuePair.Key);

                bool wasSuccessful = RecursiveReferencedProjectCheck(rootProjectKeyValuePair.Value.absoluteProjectPath, unusedProjects);

                if (!wasSuccessful)
                {
                    return null;
                }
            }

            return unusedProjects;
        }

        private bool RecursiveReferencedProjectCheck(
            string projectFilePath,
            Dictionary<Guid, (string uniqueProjectName, string absoluteProjectPath)> unusedProjects)
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
                var referencedProjectGuidString = referencedProjectMatch.Groups[2].Value;
                var referencedProjektGuid = Guid.Parse(referencedProjectGuidString);

                var searchedProject = unusedProjects.Where(p => p.Key == referencedProjektGuid).ToList();

                if (searchedProject.Count > 0)
                {
                    // Now check recursive if this project has any references.
                    bool wasSuccessful = RecursiveReferencedProjectCheck(searchedProject[0].Value.absoluteProjectPath, unusedProjects);

                    if (!wasSuccessful)
                    {
                        return false;
                    }

                    unusedProjects.Remove(searchedProject[0].Key);
                }
            }

            return true;
        }

        #endregion Private Methods
    }
}
