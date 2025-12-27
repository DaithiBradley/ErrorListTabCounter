using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using Task = System.Threading.Tasks.Task;

namespace ErrorListTabCounter
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration(
        "Error List Tab Counter",
        "Shows Errors (N) in the Error List tab",
        "1.0")]
    [ProvideAutoLoad(
        VSConstants.UICONTEXT.ShellInitialized_string,
        PackageAutoLoadFlags.BackgroundLoad)]
    [Guid(PackageGuidString)]
    public sealed class VSPackage1 : AsyncPackage

    {
        private DTE2 _dte;
        private BuildEvents _buildEvents;
        private System.Threading.Timer _timer;
        private int _lastErrorCount = -1;

        /// <summary>
        /// VSPackage1 GUID string.
        /// </summary>
        public const string PackageGuidString = "f585a34e-5c53-4154-9b95-6044c4445348";

        /// <summary>
        /// Initializes a new instance of the <see cref="VSPackage1"/> class.
        /// </summary>
        public VSPackage1()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.
        }

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        /// <param name="cancellationToken">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
        /// <param name="progress">A provider for progress updates.</param>
        /// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>
        protected override async Task InitializeAsync(
            CancellationToken cancellationToken,
            IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            HookBuildEvents();
            UpdateErrorListCaption();

            _timer = new System.Threading.Timer(_ =>
            {
                JoinableTaskFactory.RunAsync(async () =>
                {
                    await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
                    UpdateErrorListCaption();
                });
            }, null, System.TimeSpan.FromSeconds(1), System.TimeSpan.FromSeconds(1));
        }

        private void HookBuildEvents()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            _dte = (DTE2)GetService(typeof(DTE));
            Events2 events2 = (Events2)_dte.Events;
            _buildEvents = events2.BuildEvents;

            _buildEvents.OnBuildBegin += (vsBuildScope scope, vsBuildAction action) =>
    {
        ThreadHelper.ThrowIfNotOnUIThread();
        UpdateErrorListCaption();
    };

            _buildEvents.OnBuildDone += (vsBuildScope scope, vsBuildAction action) =>
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                UpdateErrorListCaption();
            };

        }

        private void UpdateErrorListCaption()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            int errorCount = GetErrorCount();

            if (errorCount == _lastErrorCount)
            {
                return;
            }

            _lastErrorCount = errorCount;

            string caption = $"Errors ({errorCount})";

            IVsUIShell uiShell = (IVsUIShell)GetService(typeof(SVsUIShell));
            if (uiShell == null)
            {
                return;
            }

            Guid errorListGuid = VSConstants.StandardToolWindows.ErrorList;
            uiShell.FindToolWindow(
                (uint)__VSFINDTOOLWIN.FTW_fForceCreate,
                ref errorListGuid,
                out IVsWindowFrame frame);

            frame?.SetProperty((int)__VSFPROPID.VSFPROPID_Caption, caption);
        }


        private int GetErrorCount()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            ErrorList errorList = _dte.ToolWindows.ErrorList;
            ErrorItems items = errorList.ErrorItems;

            int count = 0;

            for (int i = 1; i <= items.Count; i++)
            {
                ErrorItem item = items.Item(i);
                if (item.ErrorLevel == vsBuildErrorLevel.vsBuildErrorLevelHigh)
                {
                    count++;
                }
            }

            return count;
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _timer?.Dispose();
                _timer = null;
            }

            base.Dispose(disposing);
        }

        #endregion
    }
}
