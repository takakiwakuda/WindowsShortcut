using System;
using System.IO;
using System.Runtime.InteropServices;
using WindowsShortcut.Interop;

namespace WindowsShortcut.Tests;

public class TestShellLink : IDisposable
{
    public string Name { get; }
    public string? Target { get; set; }
    public string? TargetIDList { get; set; }
    public string? Arguments { get; set; }
    public string? WorkingDirectory { get; set; }
    public HotKey HotKey { get; set; }
    public WindowStyle WindowStyle { get; set; }
    public string? Description { get; set; }
    public IconLocation? IconLocation { get; set; }
    public LinkFlags LinkFlags { get; set; }

    private bool _disposed;

    public TestShellLink()
    {
        Name = Path.Combine(Path.GetTempPath(), "_Notepad.LNK");
        Target = @"%windir%\system32\notepad.exe";
        TargetIDList = NormalizeIDList(Target);
        Arguments = "document.txt";
        WorkingDirectory = "%USERPROFILE%";
        HotKey = HotKey.None;
        WindowStyle = WindowStyle.Normal;
        Description = "Text Editor";
        IconLocation = new IconLocation(@"%windir%\system32\imageres.dll", 0);
        LinkFlags = LinkFlags.HasLinkTargetIDList | LinkFlags.HasLinkInfo | LinkFlags.HasName | LinkFlags.HasRelativePath | LinkFlags.HasWorkingDirectory
                    | LinkFlags.HasArguments | LinkFlags.HasIconLocation | LinkFlags.IsUnicode | LinkFlags.HasExpString;

        CreateShortcut();
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposed)
        {
            return;
        }
        _disposed = true;

        if (disposing)
        {
            try
            {
                FileInfo fileInfo = new(Name);
                if (fileInfo.Exists)
                {
                    fileInfo.Delete();
                }
            }
            catch (Exception)
            {
                // Ignore errors
            }
        }
    }

    private static dynamic CreateInstance(string progID)
    {
        var type = Type.GetTypeFromProgID(progID) ?? throw new ArgumentException("Cannot find the specified ProgID.", nameof(progID));
        return Activator.CreateInstance(type) ?? throw new InvalidOperationException("Cannot initialize a new instance of the specified ProgID.");
    }

    private static string NormalizeIDList(string name)
    {
        string expandedName = Environment.ExpandEnvironmentVariables(name);

        int errorCode = Shell32.SHParseDisplayName(expandedName, 0, out nint ptr, 0, out ulong _);
        if (errorCode != HResults.S_OK)
        {
            throw Marshal.GetExceptionForHR(errorCode)!;
        }

        Span<char> path = stackalloc char[Kernel32.MAX_PATH];
        try
        {
            if (!Shell32.SHGetPathFromIDListW(ptr, ref MemoryMarshal.GetReference(path)))
            {
                throw new IOException($"Could not normalize the name '{name}'.");
            }
        }
        finally
        {
            Marshal.FreeCoTaskMem(ptr);
        }

        return path.Slice(0, path.IndexOf('\0')).ToString();
    }

    private void CreateShortcut()
    {
        dynamic shell = CreateInstance("WScript.Shell");
        try
        {
            dynamic shortcut = shell.CreateShortcut(Name);
            try
            {
                shortcut.TargetPath = Target;
                shortcut.Arguments = Arguments;
                shortcut.WorkingDirectory = WorkingDirectory;
                shortcut.Description = Description;
                shortcut.WindowStyle = (int)WindowStyle;

                if (HotKey.RawData > 0)
                {
                    shortcut.HotKey = HotKey.ToString().Replace("Control", "Ctrl");
                }

                if (IconLocation is not null)
                {
                    shortcut.IconLocation = IconLocation.ToString();
                }

                shortcut.Save();
            }
            finally
            {
                Marshal.FinalReleaseComObject(shortcut);
            }
        }
        finally
        {
            Marshal.FinalReleaseComObject(shell);
        }
    }
}
