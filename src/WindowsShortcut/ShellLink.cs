using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using WindowsShortcut.Interop;

namespace WindowsShortcut;

/// <summary>
/// Provides properties and methods for creating and editing shell link files.
/// </summary>
public sealed class ShellLink : IDisposable
{
    /// <summary>
    /// Gets the full path of the shell link file.
    /// </summary>
    public string Name => _name;

    /// <summary>
    /// Gets or sets the target path of the shell link file.
    /// </summary>
    public string? Target
    {
        get
        {
            ThrowIfDisposed();

            if ((LinkFlags & LinkFlags.HasLinkInfo) == 0)
            {
                return null;
            }

            _target ??= GetTargetPath(Shell32.SLGP_FLAGS.SLGP_RAWPATH);
            return _target;
        }
        set
        {
            ThrowIfDisposed();

            _shellLink.SetPath(value);
            _target = null;

            if (string.IsNullOrEmpty(value))
            {
                LinkFlags &= ~LinkFlags.HasLinkInfo;
            }
            else
            {
                LinkFlags |= LinkFlags.HasLinkInfo;
            }
        }
    }

    /// <summary>
    /// Gets or sets the item identifier for the target of the shell link file.
    /// </summary>
    public string? TargetIDList
    {
        get
        {
            ThrowIfDisposed();

            if ((LinkFlags & LinkFlags.HasLinkTargetIDList) == 0)
            {
                return null;
            }

            if (_targetIDList is null)
            {
                _shellLink.GetIDList(out nint ptr);
                try
                {
                    Span<char> path = stackalloc char[Kernel32.MAX_PATH];
                    if (!Shell32.SHGetPathFromIDListW(ptr, ref MemoryMarshal.GetReference(path)))
                    {
                        throw new IOException("Cannot retrieve item identifiers for the link target.");
                    }
                    _targetIDList = path.Slice(0, path.IndexOf('\0')).ToString();
                }
                finally
                {
                    Marshal.FreeCoTaskMem(ptr);
                }
            }
            return _targetIDList;
        }
        set
        {
            ThrowIfDisposed();

            if (value is null)
            {
                throw new ArgumentNullException(nameof(value));
            }

            if (value.Length == 0)
            {
                throw new ArgumentException("The value cannot be an empty string.", nameof(value));
            }

            int errorCode = Shell32.SHParseDisplayName(value, 0, out nint ptr, 0, out ulong _);
            if (errorCode != HResults.S_OK)
            {
                throw Marshal.GetExceptionForHR(errorCode)!;
            }

            try
            {
                _shellLink.SetIDList(ptr);
                _targetIDList = null;
                LinkFlags |= LinkFlags.HasLinkTargetIDList;
            }
            finally
            {
                Marshal.FreeCoTaskMem(ptr);
            }
        }
    }

    /// <summary>
    /// Gets or sets the command line arguments associated with the shell link.
    /// </summary>
    public string? Arguments
    {
        get
        {
            ThrowIfDisposed();

            if ((LinkFlags & LinkFlags.HasArguments) == 0)
            {
                return null;
            }

            if (_arguments is null)
            {
                Span<char> arguments = new char[Comctl32.INFOTIPSIZE];
                _shellLink.GetArguments(ref MemoryMarshal.GetReference(arguments), arguments.Length);
                _arguments = arguments.Slice(0, arguments.IndexOf('\0')).ToString();
            }
            return _arguments;
        }
        set
        {
            ThrowIfDisposed();

            _shellLink.SetArguments(value);
            _arguments = null;

            if (string.IsNullOrEmpty(value))
            {
                LinkFlags &= ~LinkFlags.HasArguments;
            }
            else
            {
                LinkFlags |= LinkFlags.HasArguments;
            }
        }
    }

    /// <summary>
    /// Gets or sets the working directory for the shell link file.
    /// </summary>
    public string? WorkingDirectory
    {
        get
        {
            ThrowIfDisposed();

            if ((LinkFlags & LinkFlags.HasWorkingDirectory) == 0)
            {
                return null;
            }

            if (_workingDirectory is null)
            {
                Span<char> workingDirectory = stackalloc char[Kernel32.MAX_PATH];
                _shellLink.GetWorkingDirectory(ref MemoryMarshal.GetReference(workingDirectory), workingDirectory.Length);
                _workingDirectory = workingDirectory.Slice(0, workingDirectory.IndexOf('\0')).ToString();
            }
            return _workingDirectory;
        }
        set
        {
            ThrowIfDisposed();

            _shellLink.SetWorkingDirectory(value);
            _workingDirectory = null;

            if (string.IsNullOrEmpty(value))
            {
                LinkFlags &= ~LinkFlags.HasWorkingDirectory;
            }
            else
            {
                LinkFlags |= LinkFlags.HasWorkingDirectory;
            }
        }
    }

    /// <summary>
    /// Gets or sets the keyboard shortcut for the shell link file.
    /// </summary>
    public HotKey HotKey
    {
        get
        {
            ThrowIfDisposed();

            if (_hotKey == -1)
            {
                _shellLink.GetHotkey(out ushort key);
                _hotKey = key;
            }
            return new HotKey(_hotKey);
        }
        set
        {
            ThrowIfDisposed();

            _shellLink.SetHotkey(value.RawData);
            _hotKey = -1;
        }
    }

    /// <summary>
    /// Gets or sets the expected window state of the target launched by the shell link.
    /// </summary>
    public WindowStyle WindowStyle
    {
        get
        {
            ThrowIfDisposed();

            if (_showCommand == 0)
            {
                _shellLink.GetShowCmd(out _showCommand);
            }
            return (WindowStyle)_showCommand;
        }
        set
        {
            ThrowIfDisposed();

            if (!Enum.IsDefined(typeof(WindowStyle), value))
            {
                throw new InvalidEnumArgumentException(nameof(value), (int)value, typeof(WindowStyle));
            }
            _shellLink.SetShowCmd((int)value);
            _showCommand = 0;
        }
    }

    /// <summary>
    /// Gets or sets the description for the shell link file.
    /// </summary>
    public string? Description
    {
        get
        {
            ThrowIfDisposed();

            if ((LinkFlags & LinkFlags.HasName) == 0)
            {
                return null;
            }

            if (_description is null)
            {
                Span<char> description = new char[Comctl32.INFOTIPSIZE];
                _shellLink.GetDescription(ref MemoryMarshal.GetReference(description), description.Length);
                _description = description.Slice(0, description.IndexOf('\0')).ToString();
            }
            return _description;
        }
        set
        {
            ThrowIfDisposed();

            _shellLink.SetDescription(value);
            _description = null;

            if (string.IsNullOrEmpty(value))
            {
                LinkFlags &= ~LinkFlags.HasName;
            }
            else
            {
                LinkFlags |= LinkFlags.HasName;
            }
        }
    }

    /// <summary>
    /// Gets or sets the icon location for the shell link file.
    /// </summary>
    public IconLocation? IconLocation
    {
        get
        {
            ThrowIfDisposed();

            if ((LinkFlags & LinkFlags.HasIconLocation) == 0)
            {
                return null;
            }

            if (_iconLocationPath is null)
            {
                Span<char> path = stackalloc char[Kernel32.MAX_PATH];
                _shellLink.GetIconLocation(ref MemoryMarshal.GetReference(path), path.Length, out _iconLocationIndex);
                _iconLocationPath = path.Slice(0, path.IndexOf('\0')).ToString();
            }
            return new IconLocation(_iconLocationPath, _iconLocationIndex);
        }
        set
        {
            ThrowIfDisposed();

            if (value is null)
            {
                _shellLink.SetIconLocation(null, 0);
                LinkFlags &= ~LinkFlags.HasIconLocation;
            }
            else
            {
                _shellLink.SetIconLocation(value.Path, value.Index);
                LinkFlags |= LinkFlags.HasIconLocation;
            }
            _iconLocationPath = null;
        }
    }

    /// <summary>
    /// Gets or sets the current option settings for the shell link.
    /// </summary>
    public LinkFlags LinkFlags
    {
        get
        {
            ThrowIfDisposed();

            if (_flags == -1)
            {
                DataList.GetFlags(out _flags);
            }
            return (LinkFlags)_flags;
        }
        set
        {
            ThrowIfDisposed();

            DataList.SetFlags((int)value);
            _flags = -1;
        }
    }

    private readonly string _name;
    private string? _target;
    private string? _targetIDList;
    private string? _arguments;
    private string? _workingDirectory;
    private int _hotKey;
    private int _showCommand;
    private string? _description;
    private string? _iconLocationPath;
    private int _iconLocationIndex;
    private int _flags;
    private Shell32.IShellLink _shellLink;

    private Shell32.IShellLinkDataList DataList => (Shell32.IShellLinkDataList)_shellLink;
    private Shell32.IPersistFile PersistFile => (Shell32.IPersistFile)_shellLink;

    /// <summary>
    /// Initializes a new instance of the <see cref="ShellLink"/> class with the specified path.
    /// </summary>
    private ShellLink(string path)
    {
        _name = path;
        _hotKey = -1;
        _flags = -1;
        _shellLink = new Shell32.IShellLink();
    }

    /// <summary>
    /// Creates a new <see cref="ShellLink"/> with the specified path.
    /// </summary>
    /// <param name="path">The full path of the shell link file to create.</param>
    /// <returns></returns>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="path"/> is <see langword="null"/>.
    /// </exception>
    public static ShellLink Create(string path)
    {
        if (path is null)
        {
            throw new ArgumentNullException(nameof(path));
        }

        return new ShellLink(Path.GetFullPath(path));
    }

    /// <summary>
    /// Opens a shell link file on the specified path.
    /// </summary>
    /// <param name="path">The full path of the shell link file to open.</param>
    /// <returns>A <see cref="ShellLink"/> on the specified path.</returns>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="path"/> is <see langword="null"/>.
    /// </exception>
    public static ShellLink Open(string path)
    {
        if (path is null)
        {
            throw new ArgumentNullException(nameof(path));
        }

        ShellLink shellLink = new(Path.GetFullPath(path));
        try
        {
            shellLink.Load();
        }
        catch (Exception)
        {
            shellLink.Dispose();
            throw;
        }

        return shellLink;
    }

    /// <summary>
    /// Releases all resources used by the <see cref="ShellLink"/>.
    /// </summary>
    public void Dispose()
    {
        if (_shellLink is not null)
        {
            Marshal.FinalReleaseComObject(_shellLink);
            _shellLink = null!;
        }
    }

    /// <summary>
    /// Reload the current object from the shell link file.
    /// </summary>
    /// <exception cref="ObjectDisposedException">
    /// The current object is disposed.
    /// </exception>
    public void Reload()
    {
        ThrowIfDisposed();
        Load();
        Refresh();
    }

    /// <summary>
    /// Saves the current shell link to the associated file.
    /// </summary>
    /// <exception cref="ObjectDisposedException">
    /// The current object is disposed.
    /// </exception>
    public void Save()
    {
        ThrowIfDisposed();
        PersistFile.Save(Name, true);
        Load();
        Refresh();
    }

    private string GetTargetPath(Shell32.SLGP_FLAGS flags)
    {
        Span<char> path = stackalloc char[Kernel32.MAX_PATH];
        _shellLink.GetPath(ref MemoryMarshal.GetReference(path), path.Length, out Kernel32.WIN32_FIND_DATA _, flags);
        return path.Slice(0, path.IndexOf('\0')).ToString();
    }

    private void Load(int mode = Ole32.STGM_READ) => PersistFile.Load(Name, mode);

    private void Refresh()
    {
        _target = null;
        _targetIDList = null;
        _arguments = null;
        _workingDirectory = null;
        _hotKey = -1;
        _showCommand = 0;
        _description = null;
        _iconLocationPath = null;
        _flags = -1;
    }

    private void ThrowIfDisposed()
    {
        if (_shellLink is null)
        {
            throw new ObjectDisposedException(typeof(ShellLink).FullName);
        }
    }
}
