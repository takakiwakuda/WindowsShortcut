#Requires -Version 7

using namespace System
using namespace System.IO
using namespace System.Text

[CmdletBinding(DefaultParameterSetName = "Path")]
param (
    [Parameter(
        Mandatory = $true,
        Position = 0,
        ParameterSetName = "Path",
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [string[]]
    $Path,

    [Parameter(
        Mandatory = $true,
        ParameterSetName = "LiteralPath",
        ValueFromPipelineByPropertyName = $true)]
    [Alias("PSPath", "LP")]
    [ValidateNotNullOrEmpty()]
    [string[]]
    $LiteralPath
)

begin {
    Import-Module -Name "$PSScriptRoot/IDList.psm1" -ErrorAction Stop

    #region Enums
    [FlagsAttribute()]
    enum CommonNetworkRelativeLinkFlags {
        ValidDevice = 0x00000001
        ValidNetType = 0x00000002
    }

    enum DriveType {
        DRIVE_UNKNOWN = 0x00000000
        DRIVE_NO_ROOT_DIR = 0x00000001
        DRIVE_REMOVABLE = 0x00000002
        DRIVE_FIXED = 0x00000003
        DRIVE_REMOTE = 0x00000004
        DRIVE_CDROM = 0x00000005
        DRIVE_RAMDISK = 0x00000006
    }

    [FlagsAttribute()]
    enum FileAttributesFlags {
        FILE_ATTRIBUTE_READONLY = 0x00000001
        FILE_ATTRIBUTE_HIDDEN = 0x00000002
        FILE_ATTRIBUTE_SYSTEM = 0x00000004
        Reserved1 = 0x00000008
        FILE_ATTRIBUTE_DIRECTORY = 0x00000010
        FILE_ATTRIBUTE_ARCHIVE = 0x00000020
        Reserved2 = 0x00000040
        FILE_ATTRIBUTE_NORMAL = 0x00000080
        FILE_ATTRIBUTE_TEMPORARY = 0x00000100
        FILE_ATTRIBUTE_SPARSE_FILE = 0x00000200
        FILE_ATTRIBUTE_REPARSE_POINT = 0x00000400
        FILE_ATTRIBUTE_COMPRESSED = 0x00000800
        FILE_ATTRIBUTE_OFFLINE = 0x00001000
        FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = 0x00002000
        FILE_ATTRIBUTE_ENCRYPTED = 0x00004000
    }

    [FlagsAttribute()]
    enum LinkFlags {
        None = 0x00000000
        HasLinkTargetIDList = 0x00000001
        HasLinkInfo = 0x00000002
        HasName = 0x00000004
        HasRelativePath = 0x00000008
        HasWorkingDirectory = 0x00000010
        HasArguments = 0x00000020
        HasIconLocation = 0x00000040
        IsUnicode = 0x00000080
        ForceNoLinkInfo = 0x00000100
        HasExpString = 0x00000200
        RunInSeparateProcess = 0x00000400
        Unused1 = 0x00000800
        HasDarwinID = 0x00001000
        RunAsUser = 0x00002000
        HasExpIcon = 0x00004000
        NoPidlAlias = 0x00008000
        Unused2 = 0x00010000
        RunWithShimLayer = 0x00020000
        ForceNoLinkTrack = 0x00040000
        EnableTargetMetadata = 0x00080000
        DisableLinkPathTracking = 0x00100000
        DisableKnownFolderTracking = 0x00200000
        DisableKnownFolderAlias = 0x00400000
        AllowLinkToLink = 0x00800000
        UnaliasOnSave = 0x01000000
        PreferEnvironmentPath = 0x02000000
        KeepLocalIDListForUNCTarget = 0x04000000
    }

    [FlagsAttribute()]
    enum LinkInfoFlags {
        None = 0x00000000
        VolumeIDAndLocalBasePath = 0x00000001
        CommonNetworkRelativeLinkAndPathSuffix = 0x00000002
    }

    enum ShowCommand {
        SW_SHOWNORMAL = 0x00000001
        SW_SHOWMAXIMIZED = 0x00000003
        SW_SHOWMINNOACTIVE = 0x00000007
    }
    #endregion

    #region Functions
    function Read-CommonNetworkRelativeLink {
        param (
            [BinaryReader]
            $Reader
        )

        $commonNetworkRelativeLink = [ordered]@{
            CommonNetworkRelativeLinkSize  = $Reader.ReadUInt32()
            CommonNetworkRelativeLinkFlags = [CommonNetworkRelativeLinkFlags]$Reader.ReadInt32()
            NetNameOffset                  = $Reader.ReadUInt32()
            DeviceNameOffset               = $Reader.ReadUInt32()
            NetworkProviderType            = $Reader.ReadUInt32()
        }

        if ($commonNetworkRelativeLink.CommonNetworkRelativeLinkFlags.HasFlag([CommonNetworkRelativeLinkFlags]::ValidNetType)) {
            $commonNetworkRelativeLink["NetName"] = [Encoding]::Default.GetString(
                $Reader.ReadBytes($commonNetworkRelativeLink.CommonNetworkRelativeLinkSize - $commonNetworkRelativeLink.NetNameOffset))
        }

        if ($commonNetworkRelativeLink.CommonNetworkRelativeLinkFlags.HasFlag([CommonNetworkRelativeLinkFlags]::ValidDevice)) {
            $commonNetworkRelativeLink["DeviceName"] = [Encoding]::Default.GetString(
                $Reader.ReadBytes($commonNetworkRelativeLink.CommonNetworkRelativeLinkSize - $commonNetworkRelativeLink.DeviceNameOffset))
        }

        [PSCustomObject]$commonNetworkRelativeLink
    }

    function Read-Header {
        param (
            [BinaryReader]
            $Reader
        )

        [PSCustomObject]@{
            HeaderSize     = $Reader.ReadUInt32()
            LinkCLSID      = [guid]::new($Reader.ReadBytes(16))
            LinkFlags      = [LinkFlags]$Reader.ReadInt32()
            FileAttributes = [FileAttributesFlags]$Reader.ReadInt32()
            CreationTime   = [datetime]::FromFileTime($Reader.ReadInt64())
            AccessTime     = [datetime]::FromFileTime($Reader.ReadInt64())
            WriteTime      = [datetime]::FromFileTime($Reader.ReadInt64())
            FileSize       = $Reader.ReadUInt32()
            IconIndex      = $Reader.ReadInt32()
            ShowCommand    = [ShowCommand]$Reader.ReadUInt32()
            HotKey         = $Reader.ReadUInt16()
            Reserved1      = $Reader.ReadInt16()
            Reserved2      = $Reader.ReadInt32()
            Reserved3      = $Reader.ReadInt32()
        }
    }

    function Read-LinkInfoHeader {
        param (
            [BinaryReader]
            $Reader
        )

        $linkInfo = [ordered]@{
            LinkInfoSize                    = $Reader.ReadUInt32()
            LinkInfoHeaderSize              = $Reader.ReadUInt32()
            LinkInfoFlags                   = [LinkInfoFlags]$Reader.ReadInt32()
            VolumeIDOffset                  = $Reader.ReadUInt32()
            LocalBasePathOffset             = $Reader.ReadUInt32()
            CommonNetworkRelativeLinkOffset = $Reader.ReadUInt32()
            CommonPathSuffixOffset          = $Reader.ReadUInt32()
        }

        if ($linkInfo.LinkInfoFlags.HasFlag([LinkInfoFlags]::VolumeIDAndLocalBasePath)) {
            $linkInfo["VolumeID"] = Read-VolumeID -Reader $Reader
            $linkInfo["LocalBasePath"] = [Encoding]::Default.GetString(
                $Reader.ReadBytes($linkInfo.CommonPathSuffixOffset - $linkInfo.LocalBasePathOffset))
        }

        if ($linkInfo.LinkInfoFlags.HasFlag([LinkInfoFlags]::CommonNetworkRelativeLinkAndPathSuffix)) {
            $linkInfo["CommonNetworkRelativeLink"] = Read-CommonNetworkRelativeLink -Reader $Reader
        }

        $linkInfo["CommonPathSuffix"] = [Encoding]::Default.GetString(
            $Reader.ReadBytes($linkInfo.LinkInfoSize - $linkInfo.CommonPathSuffixOffset))

        [PSCustomObject]$linkInfo
    }

    function Read-StringData {
        param (
            [BinaryReader]
            $Reader,

            [LinkFlags]
            $LinkFlags
        )

        $stringData = [ordered]@{}

        if ($LinkFlags.HasFlag([LinkFlags]::HasName)) {
            $stringData["NAME_STRING"] = Read-StringDataInternal -Reader $Reader
        }

        if ($LinkFlags.HasFlag([LinkFlags]::HasRelativePath)) {
            $stringData["RELATIVE_PATH"] = Read-StringDataInternal -Reader $Reader
        }

        if ($LinkFlags.HasFlag([LinkFlags]::HasWorkingDirectory)) {
            $stringData["WORKING_DIR"] = Read-StringDataInternal -Reader $Reader
        }

        if ($LinkFlags.HasFlag([LinkFlags]::HasArguments)) {
            $stringData["COMMAND_LINE_ARGUMENTS"] = Read-StringDataInternal -Reader $Reader
        }

        if ($LinkFlags.HasFlag([LinkFlags]::HasIconLocation)) {
            $stringData["ICON_LOCATION"] = Read-StringDataInternal -Reader $Reader
        }

        if ($stringData.Count -gt 0) {
            [PSCustomObject]$stringData
        }
    }

    function Read-StringDataInternal {
        param (
            [BinaryReader]
            $Reader
        )

        $count = $Reader.ReadUInt16()
        $string = [StringBuilder]::new($count)

        for ($i = 0; $i -lt $count; $i++) {
            $string.Append($Reader.ReadChar()) > $null
        }

        [PSCustomObject]@{
            CountCharacters = $count
            String          = $string
        }
    }

    function Read-TargetIDList {
        param (
            [BinaryReader]
            $Reader
        )

        $size = $Reader.ReadUInt16()
        [PSCustomObject]@{
            IDListSize = $size
            IDList     = [PSCustomObject]@{
                ItemIDList = [WindowsShortcut.IDList]::GetPathFromIDList($Reader.ReadBytes($size - 2))
                TerminalID = $Reader.ReadUInt16()
            }
        }
    }

    function Read-VolumeID {
        param (
            [BinaryReader]
            $Reader
        )

        $volumeID = [ordered]@{
            VolumeIDSize      = $Reader.ReadUInt32()
            DriveType         = [DriveType]$Reader.ReadInt32()
            DriveSerialNumber = $Reader.ReadUInt32()
            VolumeLabelOffset = $Reader.ReadUInt32()
        }

        if ($volumeID.VolumeLabelOffset -eq 0x00000014) {
            $volumeID["VolumeLabelOffsetUnicode"] = $Reader.ReadUInt32()
            $dataLength = $volumeID.VolumeIDSize - $volumeID.VolumeLabelOffsetUnicode
        } else {
            $dataLength = $volumeID.VolumeIDSize - $volumeID.VolumeLabelOffset
        }
        $volumeID["Data"] = [Encoding]::Default.GetString($Reader.ReadBytes($dataLength))

        [PSCustomObject]$volumeID
    }
    #endregion

    $streamOptions = New-Object -TypeName System.IO.FileStreamOptions -Property @{
        Mode   = [FileMode]::Open
        Access = [FileAccess]::Read
        Share  = [FileShare]::ReadWrite -bor [FileShare]::Delete
    }
}
process {
    foreach ($pathInfo in Resolve-Path @PSBoundParameters) {
        if (-not [File]::Exists($pathInfo.ProviderPath)) {
            continue
        }

        try {
            $stream = [FileStream]::new($pathInfo.ProviderPath, $streamOptions)
            $reader = [BinaryReader]::new($stream, [Encoding]::Unicode, $true)

            Read-Header -Reader $reader | Tee-Object -Variable header

            if ($header.LinkFlags.HasFlag([LinkFlags]::HasLinkTargetIDList)) {
                Read-TargetIDList -Reader $reader
            }

            if ($header.LinkFlags.HasFlag([LinkFlags]::HasLinkInfo)) {
                Read-LinkInfoHeader -Reader $reader
            }

            Read-StringData -Reader $reader -LinkFlags $header.LinkFlags
        } catch {
            $PSCmdlet.WriteError($_)
        } finally {
            ${reader}?.Dispose()
            ${stream}?.Dispose()
        }
    }
}
