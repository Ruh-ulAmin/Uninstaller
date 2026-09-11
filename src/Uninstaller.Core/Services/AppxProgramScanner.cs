using System.Diagnostics;
using System.Text.Json;
using Uninstaller.Core.Models;

namespace Uninstaller.Core.Services;

/// <summary>
/// Enumerates installed Store (AppX/MSIX) packages for the current user via
/// PowerShell's Appx module. This covers modern apps (e.g. Calculator,
/// Photos, third-party Store apps) that never appear in the classic
/// "Uninstall" registry keys.
/// </summary>
public sealed class AppxProgramScanner
{
    private sealed class AppxRecord
    {
        public string? Name { get; set; }
        public string? PackageFullName { get; set; }
        public string? Publisher { get; set; }
        public string? Version { get; set; }
        public string? InstallLocation { get; set; }
        public bool NonRemovable { get; set; }
    }

    public async Task<IReadOnlyList<InstalledProgram>> ScanAsync(CancellationToken cancellationToken = default)
    {
        // Get-AppxPackage's Version property is a [version] object, not a
        // string - ConvertTo-Json would otherwise emit it as a nested
        // {Major,Minor,Build,Revision} object and break deserialization below,
        // so it's projected through .ToString() explicitly.
        const string script =
            "Get-AppxPackage | Where-Object { -not $_.IsFramework } | " +
            "Select-Object Name,PackageFullName,Publisher,InstallLocation,NonRemovable," +
            "@{Name='Version';Expression={$_.Version.ToString()}} | " +
            "ConvertTo-Json -Compress";

        var psi = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = $"-NoProfile -NonInteractive -Command \"{script}\"",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        try
        {
            using var process = new Process { StartInfo = psi };
            process.Start();
            var stdOutTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
            var stdErrTask = process.StandardError.ReadToEndAsync(cancellationToken);
            await process.WaitForExitAsync(cancellationToken).ConfigureAwait(false);
            var stdOut = await stdOutTask.ConfigureAwait(false);
            await stdErrTask.ConfigureAwait(false);

            if (string.IsNullOrWhiteSpace(stdOut))
            {
                return Array.Empty<InstalledProgram>();
            }

            return ParseJson(stdOut);
        }
        catch
        {
            // PowerShell unavailable or Appx module missing (e.g. Server Core) -
            // simply report no Store apps rather than failing the whole scan.
            return Array.Empty<InstalledProgram>();
        }
    }

    private static IReadOnlyList<InstalledProgram> ParseJson(string json)
    {
        json = json.Trim();
        var results = new List<InstalledProgram>();

        try
        {
            if (json.StartsWith('['))
            {
                var records = JsonSerializer.Deserialize<List<AppxRecord>>(json);
                if (records is not null)
                {
                    foreach (var record in records)
                    {
                        AddIfValid(results, record);
                    }
                }
            }
            else if (json.StartsWith('{'))
            {
                var record = JsonSerializer.Deserialize<AppxRecord>(json);
                AddIfValid(results, record);
            }
        }
        catch (JsonException)
        {
            // Malformed output - return whatever we already parsed (nothing).
        }

        return results;
    }

    private static void AddIfValid(List<InstalledProgram> results, AppxRecord? record)
    {
        if (record is null || string.IsNullOrWhiteSpace(record.Name) || record.NonRemovable)
        {
            return;
        }

        results.Add(new InstalledProgram
        {
            DisplayName = record.Name!,
            DisplayVersion = record.Version,
            Publisher = record.Publisher,
            InstallLocation = record.InstallLocation,
            Source = ProgramSource.AppxPackage,
            Scope = ProgramScope.User,
            Architecture = ProgramArchitecture.Unknown,
            PackageFullName = record.PackageFullName
        });
    }
}
