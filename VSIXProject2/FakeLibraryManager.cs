using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Runtime.InteropServices;

namespace VSIXProject2
{
    internal class FakeLibraryManager : IVsLibraryMgr
    {
        public int GetCount(out uint pnCount)
        {
            pnCount = 0;
            return VSConstants.S_OK;
        }

        public int GetLibraryAt(uint nLibIndex, out IVsLibrary ppLibrary)
        {
            ppLibrary = null;
            return VSConstants.S_OK;
        }

        public int GetNameAt(uint nLibIndex, IntPtr pszName)
        {
            pszName = Marshal.StringToBSTR("");
            return VSConstants.S_OK;
        }

        public int ToggleCheckAt(uint nLibIndex)
        {
            return VSConstants.S_OK;
        }

        public int GetCheckAt(uint nLibIndex, LIB_CHECKSTATE[] pstate)
        {
            pstate = null;
            return VSConstants.S_OK;
        }

        public int SetLibraryGroupEnabled(LIB_PERSISTTYPE lpt, int fEnable)
        {
            return VSConstants.S_OK;
        }
    }
}