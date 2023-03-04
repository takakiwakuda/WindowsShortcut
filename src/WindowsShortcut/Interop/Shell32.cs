using System;
using System.Runtime.InteropServices;

namespace WindowsShortcut.Interop;

internal static partial class Shell32
{
    internal enum SLGP_FLAGS
    {
        SLGP_SHORTPATH = 0x1,
        SLGP_UNCPRIORITY = 0x2,
        SLGP_RAWPATH = 0x4,
        SLGP_RELATIVEPRIORITY = 0x8
    }

    [ComImport]
    [Guid("0000010b-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IPersistFile
    {
        void GetClassID(out Guid GetClassID);
        [PreserveSig]
        int IsDirty();
        void Load([MarshalAs(UnmanagedType.LPWStr)] string pszFileName, int dwMode);
        void Save([MarshalAs(UnmanagedType.LPWStr)] string? pszFileName, [MarshalAs(UnmanagedType.Bool)] bool fRemember);
        void SaveCompleted([MarshalAs(UnmanagedType.LPWStr)] string pszFileName);
        void GetCurFile([MarshalAs(UnmanagedType.LPWStr)] out string? ppszFileName);
    }

    [ComImport]
    [Guid("000214F9-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [CoClass(typeof(ShellLinkClass))]
    internal interface IShellLink
    {
        void GetPath(ref char pszFile, int cch, out Kernel32.WIN32_FIND_DATA pfd, SLGP_FLAGS fFlags);
        void GetIDList(out nint ppidl);
        void SetIDList(nint pidl);
        void GetDescription(ref char pszName, int cch);
        void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string? pszName);
        void GetWorkingDirectory(ref char pszDir, int cch);
        void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string? pszDir);
        void GetArguments(ref char pszArgs, int cch);
        void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string? pszArgs);
        void GetHotkey(out ushort pwHotkey);
        void SetHotkey(ushort wHotkey);
        void GetShowCmd(out int piShowCmd);
        void SetShowCmd(int iShowCmd);
        void GetIconLocation(ref char pszIconPath, int cch, out int piIcon);
        void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string? pszIconPath, int iIcon);
        void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, uint dwReserved);
        void Resolve(nint hwnd, int fFlags);
        void SetPath([MarshalAs(UnmanagedType.LPWStr)] string? pszFile);
    }

    [ComImport]
    [Guid("45E2B4AE-B1C3-11D0-B92F-00A0C90312E1")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IShellLinkDataList
    {
        void AddDataBlock(nint pDataBlock);
        void CopyDataBlock(uint dwSig, out nint ppDataBlock);
        void RemoveDataBlock(uint dwSig);
        void GetFlags(out int pdwFlags);
        void SetFlags(int dwFlags);
    }

    [ComImport]
    [Guid("00021401-0000-0000-C000-000000000046")]
    internal class ShellLinkClass
    {
    }

#if NET7_0_OR_GREATER
    [LibraryImport(nameof(Shell32), SetLastError = true, StringMarshalling = StringMarshalling.Utf16)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool SHGetPathFromIDListW(nint pidl, ref char pszPath);

    [LibraryImport(nameof(Shell32), SetLastError = true, StringMarshalling = StringMarshalling.Utf16)]
    internal static partial int SHParseDisplayName(string pszName, nint pbc, out nint ppidl, ulong sfgaoIn, out ulong psfgaoOut);
#else
    [DllImport(nameof(Shell32), SetLastError = true, CharSet = CharSet.Unicode)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool SHGetPathFromIDListW(nint pidl, ref char pszPath);

    [DllImport(nameof(Shell32), SetLastError = true, CharSet = CharSet.Unicode)]
    internal static extern int SHParseDisplayName(string pszName, nint pbc, out nint ppidl, ulong sfgaoIn, out ulong psfgaoOut);
#endif
}
