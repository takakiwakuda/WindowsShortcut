[CmdletBinding()]
param (
    [Parameter()]
    [ValidateSet("Debug", "Release")]
    [string]
    $Configuration = (property Configuration Release),

    [Parameter()]
    [ValidateSet("net462", "net7.0")]
    [string]
    $Framework = "net462"
)

<#
.SYNOPSIS
    Build WindowsShortcut assembly
#>
task Build @{
    Inputs  = {
        Get-ChildItem src/WindowsShortcut/*.cs, src/WindowsShortcut/Interop/*.cs, src/WindowsShortcut/WindowsShortcut.csproj
    }
    Outputs = "src/WindowsShortcut/bin/$Configuration/$Framework/win-x64/WindowsShortcut.dll"
    Jobs    = {
        exec { dotnet publish -f $Framework -c $Configuration src/WindowsShortcut }
    }
}

<#
.SYNOPSIS
    Build WindowsShortcut.Tests assembly
#>
task BuildTests @{
    Inputs  = {
        Get-ChildItem tests/WindowsShortcut.Tests/*.cs, tests/WindowsShortcut.Tests/WindowsShortcut.Tests.csproj
    }
    Outputs = "tests/WindowsShortcut.Tests/bin/$Configuration/$Framework/win-x64/WindowsShortcut.Tests.dll"
    Jobs    = {
        exec { dotnet build -f $Framework -c $Configuration tests/WindowsShortcut.Tests }
    }
}

<#
.SYNOPSIS
    Run WindowsShortcut testing
#>
task RunTests BuildTests, {
    exec {
        dotnet test --no-build -f $Framework -c $Configuration tests/WindowsShortcut.Tests
    }
}

task . RunTests
