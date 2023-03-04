using System;
using System.Runtime.InteropServices;

namespace WindowsShortcut.Interop;

internal static class Kernel32
{
    [StructLayout(LayoutKind.Sequential)]
    internal struct FILETIME
    {
        internal uint dwLowDateTime;
        internal uint dwHighDateTime;

        internal long ToTicks() => ((long)dwHighDateTime << 32) + dwLowDateTime;
        internal DateTime ToDateTime() => DateTime.FromFileTime(ToTicks());
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    internal struct WIN32_FIND_DATA
    {
        internal int dwFileAttributes;
        internal FILETIME ftCreationTime;
        internal FILETIME ftLastAccessTime;
        internal FILETIME ftLastWriteTime;
        internal uint nFileSizeHigh;
        internal uint nFileSizeLow;
        internal uint dwReserved0;
        internal uint dwReserved1;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = MAX_PATH)]
        internal string cFileName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 14)]
        internal string cAlternateFileName;
    }

    internal const int MAX_PATH = 260;
}
