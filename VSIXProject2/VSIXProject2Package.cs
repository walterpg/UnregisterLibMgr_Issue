using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using Task = System.Threading.Tasks.Task;

namespace VSIXProject2
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
    [Guid(VSIXProject2Package.PackageGuidString)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionHasMultipleProjects_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionHasSingleProject_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    public sealed class VSIXProject2Package : AsyncPackage
    {
        /// <summary>
        /// VSIXProject2Package GUID string.
        /// </summary>
        public const string PackageGuidString = "b21819e3-c975-4a39-a132-a8baa7e8d37e";

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        /// <param name="cancellationToken">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
        /// <param name="progress">A provider for progress updates.</param>
        /// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            Cookie = VSConstants.VSCOOKIE_NIL;

            // When initialized asynchronously, the current thread may be a background thread at this point.
            // Do any initialization that requires the UI thread after switching to the UI thread.
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            if (!(await this.GetServiceAsync(typeof(SVsObjectManager)) is IVsObjectManager objectManager))
            {
                Debug.Fail("");
                return;
            }
            Guid guid = Guid.Parse(PackageGuidString);
            if (ErrorHandler.Succeeded(objectManager.RegisterLibMgr(ref guid, new FakeLibraryManager(), out uint cookie)))
            {
                Cookie = cookie;
                await JoinableTaskFactory.RunAsync(async () =>
                {
                    await Task.Yield();
                    MessageBox.Show("Success.", nameof(IVsObjectManager.RegisterLibMgr)+$"() Cookie=0x{cookie:x}");
                    return VSConstants.S_OK;
                });
            }
            await Command1.InitializeAsync(this);
        }

        public uint Cookie { get; set; }

        protected override void Dispose(bool disposing)
        {
            if (Cookie != VSConstants.VSCOOKIE_NIL && !Zombied)
            {
                uint cookie = Cookie;
                _ = JoinableTaskFactory.Run(async () =>
                {
                    await JoinableTaskFactory.SwitchToMainThreadAsync();
                    if (await GetServiceAsync(typeof(SVsObjectManager)) is IVsObjectManager objectManager)
                    {
                        string message = "Success.";
                        if (VSConstants.E_INVALIDARG == objectManager.UnregisterLibMgr(cookie))
                        {
                            Debug.WriteLine("UnregisterLibMgr failed.");
                            message = nameof(VSConstants.E_INVALIDARG) +
                                $". Test machine shows 'Access vialoation reading location 0x00000000{cookie+0x20:x}'.";
                        }
                        await JoinableTaskFactory.RunAsync(async () =>
                        {
                            await Task.Yield();
                            MessageBox.Show(message, nameof(IVsObjectManager.UnregisterLibMgr) +$"(0x{cookie:x})");
                            return VSConstants.S_OK;
                        });
                    }
                    return VSConstants.S_OK;
                });
            }
            Cookie = VSConstants.VSCOOKIE_NIL;
            base.Dispose(disposing);
        }

        #endregion
    }
}
