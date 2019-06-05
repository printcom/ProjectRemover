using System;
using System.Collections.ObjectModel;
using System.Windows.Input;
using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.CommandWpf;
using ProjectRemover.Package.Classes;

namespace ProjectRemover.Package.Windows
{
    public class RemoveProjectsViewModel : ViewModelBase
    {
        private ICommand _cmdCloseWindow;
        private ICommand _cmdApprove;

        #region Getter / Setter

        public Action CloseWindowAction { get; set; }

        public bool IsCanceled { get; set; } = true;

        #endregion Getter / Setter

        #region Bindings

        public ObservableCollection<RemovableProject> RemovableProjects { get; set; } = new ObservableCollection<RemovableProject>();

        #endregion Bindings

        #region Commands

        public ICommand CmdCloseWindow
        {
            get { return _cmdCloseWindow ?? (_cmdCloseWindow = new RelayCommand(CloseWindow)); }
        }

        public ICommand CmdApprove
        {
            get { return _cmdApprove ?? (_cmdApprove = new RelayCommand(Approve)); }
        }

        #endregion Commands

        #region Private Methods

        private void Approve()
        {
            IsCanceled = false;
            CloseWindowAction?.Invoke();
        }

        private void CloseWindow()
        {
            IsCanceled = true;
            CloseWindowAction?.Invoke();
        }

        #endregion Private Methods
    }
}