using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace StackTraceExtension.Helpers
{
    static class ProjectUtilities
    {
        private static IServiceProvider _serviceProvider;

        public static void SetServiceProvider(IServiceProvider provider)
        {
            _serviceProvider = provider;
        }

        static public IList<IVsProject> GetProjectsOfCurrentSelections()
        {
            List<IVsProject> results = new List<IVsProject>();

            int hr = VSConstants.S_OK;
            var selectionMonitor = _serviceProvider.GetService(typeof(SVsShellMonitorSelection)) as IVsMonitorSelection;
            if (selectionMonitor == null)
            {
                Debug.Fail("Failed to get SVsShellMonitorSelection service.");
                return results;
            }

            IntPtr hierarchyPtr = IntPtr.Zero;
            uint itemID = 0;
            IVsMultiItemSelect multiSelect = null;
            IntPtr containerPtr = IntPtr.Zero;
            hr = selectionMonitor.GetCurrentSelection(out hierarchyPtr, out itemID, out multiSelect, out containerPtr);
            if (IntPtr.Zero != containerPtr)
            {
                Marshal.Release(containerPtr);
                containerPtr = IntPtr.Zero;
            }
            Debug.Assert(hr == VSConstants.S_OK, "GetCurrentSelection failed.");

            if (itemID == (uint)VSConstants.VSITEMID.Selection)
            {
                uint itemCount = 0;
                int fSingleHierarchy = 0;
                hr = multiSelect.GetSelectionInfo(out itemCount, out fSingleHierarchy);
                Debug.Assert(hr == VSConstants.S_OK, "GetSelectionInfo failed.");

                VSITEMSELECTION[] items = new VSITEMSELECTION[itemCount];
                hr = multiSelect.GetSelectedItems(0, itemCount, items);
                Debug.Assert(hr == VSConstants.S_OK, "GetSelectedItems failed.");

                foreach (VSITEMSELECTION item in items)
                {
                    IVsProject project = GetProjectOfItem(item.pHier, item.itemid);
                    if (!results.Contains(project))
                    {
                        results.Add(project);
                    }
                }
            }
            else
            {
                // case where no visible project is open (single file)
                if (hierarchyPtr != IntPtr.Zero)
                {
                    IVsHierarchy hierarchy = (IVsHierarchy)Marshal.GetUniqueObjectForIUnknown(hierarchyPtr);
                    results.Add(GetProjectOfItem(hierarchy, itemID));
                }
            }

            return results;
        }

        private static IVsProject GetProjectOfItem(IVsHierarchy hierarchy, uint itemID)
        {
            return (IVsProject)hierarchy;
        }

        public static string GetProjectFilePath(IVsProject project)
        {
            string path = string.Empty;
            int hr = project.GetMkDocument((uint)VSConstants.VSITEMID.Root, out path);
            Debug.Assert(hr == VSConstants.S_OK || hr == VSConstants.E_NOTIMPL, "GetMkDocument failed for project.");

            return path;
        }

        public static string GetUniqueProjectNameFromFile(string projectFile)
        {
            IVsProject project = GetProjectByFileName(projectFile);

            if (project != null)
            {
                return GetUniqueUIName(project);
            }

            return null;
        }

        public static string GetUniqueUIName(IVsProject project)
        {
            var solution = _serviceProvider.GetService(typeof(SVsSolution)) as IVsSolution3;
            if (solution == null)
            {
                Debug.Fail("Failed to get SVsSolution service.");
                return null;
            }

            string name = null;
            int hr = solution.GetUniqueUINameOfProject((IVsHierarchy)project, out name);
            Debug.Assert(hr == VSConstants.S_OK, "GetUniqueUINameOfProject failed.");
            return name;
        }

        public static IEnumerable<IVsProject> LoadedProjects
        {
            get
            {
                var solution = _serviceProvider.GetService(typeof(SVsSolution)) as IVsSolution;
                if (solution == null)
                {
                    Debug.Fail("Failed to get SVsSolution service.");
                    yield break;
                }

                IEnumHierarchies enumerator = null;
                Guid guid = Guid.Empty;
                solution.GetProjectEnum((uint)__VSENUMPROJFLAGS.EPF_LOADEDINSOLUTION, ref guid, out enumerator);
                IVsHierarchy[] hierarchy = new IVsHierarchy[1] { null };
                uint fetched = 0;
                for (enumerator.Reset(); enumerator.Next(1, hierarchy, out fetched) == VSConstants.S_OK && fetched == 1; /*nothing*/)
                {
                    yield return (IVsProject)hierarchy[0];
                }
            }
        }

        public static IVsProject GetProjectByFileName(string projectFile)
        {
            return LoadedProjects.FirstOrDefault(
                p => string.Compare(projectFile, GetProjectFilePath(p), StringComparison.OrdinalIgnoreCase) == 0);
        }

        public static IEnumerable<string> AllItemsInProject(IVsProject project)
        {
            if (project == null)
            {
                throw new ArgumentNullException("project");
            }

            string projectDir = Path.GetDirectoryName(ProjectUtilities.GetProjectFilePath(project));
            IVsHierarchy hierarchy = project as IVsHierarchy;

            return
                ChildrenOf(hierarchy, VSConstants.VSITEMID.Root)
                .Select(
                    id =>
                    {
                        string name = null;
                        project.GetMkDocument((uint)id, out name);
                        if (name != null && name.Length > 0 && !Path.IsPathRooted(name))
                        {
                            name = AbsolutePathFromRelative(name, projectDir);
                        }
                        return name;
                    })
                .Where(File.Exists);
        }

        /// <summary>
        /// Transforms a relative path to an absolute one based on a specified base folder.
        /// </summary>
        static public string AbsolutePathFromRelative(string relativePath, string baseFolderForDerelativization)
        {
            if (relativePath == null)
            {
                throw new ArgumentNullException("relativePath");
            }
            if (baseFolderForDerelativization == null)
            {
                throw new ArgumentNullException("baseFolderForDerelativization");
            }
            if (Path.IsPathRooted(relativePath))
            {
                throw new ArgumentException("Path Not Relative", "relativePath");
            }
            if (!Path.IsPathRooted(baseFolderForDerelativization))
            {
                throw new ArgumentException("Base Folder Must Be Rooted", "baseFolderForDerelativization");
            }

            StringBuilder result = new StringBuilder(baseFolderForDerelativization);

            if (result[result.Length - 1] != Path.DirectorySeparatorChar)
            {
                result.Append(Path.DirectorySeparatorChar);
            }

            int spanStart = 0;

            while (spanStart < relativePath.Length)
            {
                int spanStop = relativePath.IndexOf(Path.DirectorySeparatorChar, spanStart);

                if (spanStop == -1)
                {
                    spanStop = relativePath.Length;
                }

                string span = relativePath.Substring(spanStart, spanStop - spanStart);

                if (span == "..")
                {
                    // The result string should end with a directory separator at this point.  We
                    // want to search for the one previous to that, which is why we subtract 2.
                    int previousSeparator;
                    if (result.Length < 2 || (previousSeparator = result.ToString().LastIndexOf(Path.DirectorySeparatorChar, result.Length - 2)) == -1)
                    {
                        throw new ArgumentException("Back Too Far");
                    }
                    result.Remove(previousSeparator + 1, result.Length - previousSeparator - 1);
                }
                else if (span != ".")
                {
                    // Ignore "." because it means the current direcotry
                    result.Append(span);

                    if (spanStop < relativePath.Length)
                    {
                        result.Append(Path.DirectorySeparatorChar);
                    }

                }

                spanStart = spanStop + 1;
            }

            return result.ToString();
        }

        private static List<VSConstants.VSITEMID> ChildrenOf(IVsHierarchy hierarchy, VSConstants.VSITEMID rootID)
        {
            var result = new List<VSConstants.VSITEMID>();

            for (VSConstants.VSITEMID itemID = FirstChild(hierarchy, rootID); itemID != VSConstants.VSITEMID.Nil; itemID = NextSibling(hierarchy, itemID))
            {
                result.Add(itemID);
                result.AddRange(ChildrenOf(hierarchy, itemID));
            }

            return result;
        }

        private static VSConstants.VSITEMID FirstChild(IVsHierarchy hierarchy, VSConstants.VSITEMID rootID)
        {
            object childIDObj = null;
            hierarchy.GetProperty((uint)rootID, (int)__VSHPROPID.VSHPROPID_FirstChild, out childIDObj);
            if (childIDObj != null)
            {
                return (VSConstants.VSITEMID)(int)childIDObj;
            }

            return VSConstants.VSITEMID.Nil;
        }

        private static VSConstants.VSITEMID NextSibling(IVsHierarchy hierarchy, VSConstants.VSITEMID firstID)
        {
            object siblingIDObj = null;
            hierarchy.GetProperty((uint)firstID, (int)__VSHPROPID.VSHPROPID_NextSibling, out siblingIDObj);
            if (siblingIDObj != null)
            {
                return (VSConstants.VSITEMID)(int)siblingIDObj;
            }

            return VSConstants.VSITEMID.Nil;
        }

        //static public bool IsMSBuildProject(IVsProject project)
        //{
        //    return ProjectCollection.GlobalProjectCollection.GetLoadedProjects(GetProjectFilePath(project)).Any();
        //}
    }
}
