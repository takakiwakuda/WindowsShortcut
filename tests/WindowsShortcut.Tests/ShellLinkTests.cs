using System.IO;
using Xunit;

namespace WindowsShortcut.Tests;

public class ShellLinkTests : IClassFixture<TestShellLink>
{
    private readonly TestShellLink _testShellLink;

    public ShellLinkTests(TestShellLink testShellLink)
    {
        _testShellLink = testShellLink;
    }

    [Fact]
    public void CreateEmptyShellLink()
    {
        using ShellLink shellLink = ShellLink.Create("_does_not_exist.LNK");

        Assert.Null(shellLink.Target);
        Assert.Null(shellLink.TargetIDList);
        Assert.Null(shellLink.Arguments);
        Assert.Null(shellLink.WorkingDirectory);
        Assert.Null(shellLink.Description);
        Assert.Null(shellLink.IconLocation);
        Assert.Equal(HotKey.None.RawData, shellLink.HotKey.RawData);
        Assert.Equal(WindowStyle.Normal, shellLink.WindowStyle);
        Assert.Equal(LinkFlags.None, shellLink.LinkFlags);
    }

    [Fact]
    public void CreateAdnSetProperties()
    {
        using ShellLink shellLink = ShellLink.Create(Path.Combine(Path.GetTempPath(), "_does_not_exist.LNK"));
        CopyProperties(_testShellLink, shellLink);

        try
        {
            shellLink.Save();
            AssertProperties(_testShellLink, shellLink);
        }
        finally
        {
            FileInfo fileInfo = new(shellLink.Name);
            if (fileInfo.Exists)
            {
                fileInfo.Delete();
            }
        }
    }

    [Fact]
    public void OpenExistingShellLink()
    {
        using ShellLink shellLink = ShellLink.Open(_testShellLink.Name);
        AssertProperties(_testShellLink, shellLink);
    }

    private static void AssertProperties(TestShellLink expected, ShellLink actual)
    {
        Assert.Equal(expected.Target, actual.Target);
        Assert.Equal(expected.TargetIDList, actual.TargetIDList);
        Assert.Equal(expected.Arguments, actual.Arguments);
        Assert.Equal(expected.WorkingDirectory, actual.WorkingDirectory);
        Assert.Equal(expected.HotKey.RawData, actual.HotKey.RawData);
        Assert.Equal(expected.WindowStyle, actual.WindowStyle);
        Assert.Equal(expected.Description, actual.Description);
        Assert.Equal(expected.IconLocation.Path, actual.IconLocation.Path);
        Assert.Equal(expected.IconLocation.Index, actual.IconLocation.Index);
        Assert.Equal(expected.LinkFlags, actual.LinkFlags);
    }

    private static void CopyProperties(TestShellLink source, ShellLink destination)
    {
        destination.Target = source.Target;
        destination.Arguments = source.Arguments;
        destination.WorkingDirectory = source.WorkingDirectory;
        destination.HotKey = new HotKey(source.HotKey.RawData);
        destination.WindowStyle = source.WindowStyle;
        destination.Description = source.Description;
        destination.IconLocation = IconLocation.Parse(source.IconLocation.ToString());
    }
}
