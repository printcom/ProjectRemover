namespace ProjectRemover.Package.Windows
{
    /// <summary>
    /// Interaction logic for RemoveProjectsWindow.xaml
    /// </summary>
    public partial class RemoveProjectsWindow 
    {
        public RemoveProjectsWindow()
        {
            InitializeComponent();
            ViewModel = new RemoveProjectsViewModel
            {
                CloseWindowAction = Close
            };

            DataContext = ViewModel;
        }

        public RemoveProjectsViewModel ViewModel { get; }
    }
}
