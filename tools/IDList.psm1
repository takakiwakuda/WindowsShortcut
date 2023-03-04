Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace WindowsShortcut;

public static class IDList
{
    public static string GetPathFromIDList(byte[] idList)
    {
        GCHandle handle = GCHandle.Alloc(idList, GCHandleType.Pinned);
        Span<char> path = stackalloc char[MAX_PATH];

        try
        {
            nint ptr = handle.AddrOfPinnedObject();
            if (!SHGetPathFromIDListW(ptr, ref MemoryMarshal.GetReference(path)))
            {
                int errorCode = Marshal.GetHRForLastWin32Error();
                throw new IOException($"Cannot retrieve a path from the specified ID list.", errorCode);
            }
        }
        finally
        {
            handle.Free();
        }

        return path[..path.IndexOf('\0')].ToString();
    }

    private const int MAX_PATH = 260;

    [DllImport("shell32", SetLastError = true, CharSet = CharSet.Unicode)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool SHGetPathFromIDListW(nint pidl, ref char pszPath);
}
'@
