using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.Threading;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using Microsoft.VisualStudio;
using EnvDTE80;
using StrongNameKeyGenerator.ViewModel;
using System.Linq;
using EnvDTE;
using System.Collections.Generic;

namespace StrongNameKeyGenerator
{

    internal sealed class SnCommand
    {
        private readonly Package _package;
        private readonly DTE2 _dte;
        private string _projectDirectory;
        private string _strongNameKey;
        public string _projectName;
        public string _filePath;

        private SnCommand(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException(nameof(package));
            }

            _package = package;
            _dte = (DTE2)ServiceProvider.GetService(typeof(DTE));


            var commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            var menuCommandID = new CommandID(PackageGuids.guidSnCommandPackageCmdSet, PackageIds.SnCommandId);
            var oleMenuCommand = new OleMenuCommand(this.ExecuteCommand, menuCommandID);
            oleMenuCommand.BeforeQueryStatus += BeforeQueryStatus;
            commandService.AddCommand(oleMenuCommand);



        }

        public static SnCommand Instance { get; private set; }

        private IServiceProvider ServiceProvider { get { return _package; } }



        public static void Initialize(Package package)
        {
            Instance = new SnCommand(package);
        }

        private void BeforeQueryStatus(object sender, EventArgs e)
        {
            var button = sender as OleMenuCommand;
            if (button != null)
            {
                button.Enabled = false;
                IVsHierarchy hierarchy = null;
                var itemid = VSConstants.VSITEMID_NIL;

                if (IsSelectedItemSingle(out hierarchy, out itemid))
                {

                    string itemFullPath = null;
                    ((IVsProject)hierarchy).GetMkDocument(itemid, out itemFullPath);
                    var transformFileInfo = new FileInfo(itemFullPath);

                    _projectDirectory = transformFileInfo.DirectoryName;
                    _projectName = transformFileInfo.Directory.Name;
                    _strongNameKey = Path.GetFileName(transformFileInfo.FullName.ToString());

                    var isStrongNameFile = transformFileInfo.Name.Contains(@".snk");

                    if (isStrongNameFile)
                    {
                        button.Enabled = true;
                    }
                }
            }
        }

        private static bool IsSelectedItemSingle(out IVsHierarchy hierarchy, out uint pitemid)
        {
            hierarchy = null;
            pitemid = VSConstants.VSITEMID_NIL;
            var hr = VSConstants.S_OK;
            var monitorSelection = Package.GetGlobalService(typeof(SVsShellMonitorSelection)) as IVsMonitorSelection;
            var solution = Package.GetGlobalService(typeof(SVsSolution)) as IVsSolution;

            if (monitorSelection == null || solution == null)
            {
                return false;
            }

            IVsMultiItemSelect ppMIS = null;
            var hierarchyPtr = IntPtr.Zero;
            var ppSC = IntPtr.Zero;

            try
            {
               
                hr = monitorSelection.GetCurrentSelection(out hierarchyPtr, out pitemid, out ppMIS, out ppSC);
                if (ErrorHandler.Failed(hr) || hierarchyPtr == IntPtr.Zero || pitemid == VSConstants.VSITEMID_NIL)
                {
                    return false;
                }
                if (ppMIS != null)
                {
                    return false;
                }
                if (pitemid == VSConstants.VSITEMID_ROOT)
                {
                    return false;
                }

                hierarchy = System.Runtime.InteropServices.Marshal.GetObjectForIUnknown(hierarchyPtr) as IVsHierarchy;

                if (hierarchy == null)
                {
                    return false;
                }
                var guidProjectID = Guid.Empty;

                if (ErrorHandler.Failed(solution.GetGuidOfProject(hierarchy, out guidProjectID)))
                {
                    return false;
                }
                return true;
            }
            finally
            {
                if (ppSC != IntPtr.Zero)
                {
                    System.Runtime.InteropServices.Marshal.Release(ppSC);
                }
                if (hierarchyPtr != IntPtr.Zero)
                {
                    System.Runtime.InteropServices.Marshal.Release(hierarchyPtr);
                }
            }
        }

        private void ExecuteCommand(object sender, EventArgs e)
        {

            try
            {
                var shell = (IVsShell)ServiceProvider.GetService(typeof(SVsShell));
                object root;

                if (shell.GetProperty((int)__VSSPROPID.VSSPROPID_VirtualRegistryRoot, out root) == VSConstants.S_OK)
                {
                    var appData = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
                    var version = GetNumbers(Path.GetFileName(root.ToString()));
                    _filePath = Path.Combine(appData, "Microsoft Visual Studio " + version + "\\Common7\\Tools\\VsDevCmd.bat");
                    Logger.Log("Strong Name Key path: " + _filePath);
                }


                var vm = new StrongNameViewModel
                {
                    DTE = _dte,
                    ProjectDirectory = _projectDirectory,
                    StrongName = _strongNameKey,
                    OutFileName = _projectName,
                    FilePath = _filePath
                };
                vm.StrongNamekeyGeneration();
            }
            catch (Exception ex)
            {
                Logger.Log("Error occurred: " + ex.Message);
            }
        }


        private static string GetNumbers(string input)
        {
            return Regex.Replace(input, "[^0-9.]", "");
        }


        public List<EnvDTE.ProjectItem> GetProjectItemsRecursively(EnvDTE.ProjectItems items)
        {
            var ret = new List<EnvDTE.ProjectItem>();
            if (items == null) return ret;
            foreach (EnvDTE.ProjectItem item in items)
            {
                ret.Add(item);
                ret.AddRange(GetProjectItemsRecursively(item.ProjectItems));
            }
            return ret;
        }


    }

}
