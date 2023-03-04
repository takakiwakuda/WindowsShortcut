﻿using System;
using System.ComponentModel;

namespace WindowsShortcut;

/// <summary>
/// Specifies options settings for a shell link.
/// </summary>
[Flags]
public enum LinkFlags
{
    None = 0x00000000,
    HasLinkTargetIDList = 0x00000001,
    HasLinkInfo = 0x00000002,
    HasName = 0x00000004,
    HasRelativePath = 0x00000008,
    HasWorkingDirectory = 0x00000010,
    HasArguments = 0x00000020,
    HasIconLocation = 0x00000040,
    IsUnicode = 0x00000080,
    ForceNoLinkInfo = 0x00000100,
    HasExpString = 0x00000200,
    RunInSeparateProcess = 0x00000400,
    [EditorBrowsable(EditorBrowsableState.Never)]
    Unused1 = 0x00000800,
    HasDarwinID = 0x00001000,
    RunAsUser = 0x00002000,
    HasExpIcon = 0x00004000,
    NoPidlAlias = 0x00008000,
    [EditorBrowsable(EditorBrowsableState.Never)]
    Unused2 = 0x00010000,
    RunWithShimLayer = 0x00020000,
    ForceNoLinkTrack = 0x00040000,
    EnableTargetMetadata = 0x00080000,
    DisableLinkPathTracking = 0x00100000,
    DisableKnownFolderTracking = 0x00200000,
    DisableKnownFolderAlias = 0x00400000,
    AllowLinkToLink = 0x00800000,
    UnaliasOnSave = 0x01000000,
    PreferEnvironmentPath = 0x02000000,
    KeepLocalIDListForUNCTarget = 0x04000000
}
