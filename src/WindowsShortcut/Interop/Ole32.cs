namespace WindowsShortcut.Interop;

internal static class Ole32
{
    internal const int STGM_READ = 0x00000000;
    internal const int STGM_WRITE = 0x00000001;
    internal const int STGM_READWRITE = 0x00000002;
    internal const int STGM_SHARE_DENY_NONE = 0x00000040;
    internal const int STGM_SHARE_DENY_READ = 0x00000030;
    internal const int STGM_SHARE_DENY_WRITE = 0x00000020;
    internal const int STGM_SHARE_EXCLUSIVE = 0x00000010;
    internal const int STGM_PRIORITY = 0x00040000;
    internal const int STGM_CREATE = 0x00001000;
    internal const int STGM_CONVERT = 0x00020000;
    internal const int STGM_FAILIFTHERE = 0x00000000;
    internal const int STGM_DIRECT = 0x00000000;
    internal const int STGM_TRANSACTED = 0x00010000;
    internal const int STGM_NOSCRATCH = 0x00100000;
    internal const int STGM_NOSNAPSHOT = 0x00200000;
    internal const int STGM_SIMPLE = 0x08000000;
    internal const int STGM_DIRECT_SWMR = 0x00400000;
    internal const int STGM_DELETEONRELEASE = 0x04000000;
}
