using System.Diagnostics;
using System.Reflection;
using System.Text.Json;
using Microsoft.Win32;

namespace CertificateAutomater;

public sealed class UpdateInfo
{
    public required Version LatestVersion { get; init; }
    public required string TagName { get; init; }
    public required string ReleaseName { get; init; }
    public required string ReleaseNotes { get; init; }
    public required string InstallerDownloadUrl { get; init; }
    public required string InstallerFileName { get; init; }
    public required string ReleasePageUrl { get; init; }
}

public static class UpdateChecker
{
    private const string GitHubOwner = "wizexplorer";
    private const string GitHubRepo = "CertificateAutomater";

    // should match the asset on GitHub Releases
    private const string InstallerAssetName = "CertificateAutomaterSetup.exe";

    public static async Task<UpdateInfo?> CheckForUpdateAsync()
    {
        Version currentVersion = GetCurrentVersion();

        using HttpClient httpClient = new HttpClient();

        httpClient.DefaultRequestHeaders.UserAgent.ParseAdd(
            "CertificateAutomater-UpdateChecker"
        );

        string apiUrl =
            $"https://api.github.com/repos/{GitHubOwner}/{GitHubRepo}/releases/latest";

        using HttpResponseMessage response = await httpClient.GetAsync(apiUrl);

        if (!response.IsSuccessStatusCode)
        {
            throw new Exception(
                $"GitHub update check failed. Status: {(int)response.StatusCode} {response.ReasonPhrase}"
            );
        }

        string json = await response.Content.ReadAsStringAsync();

        using JsonDocument document = JsonDocument.Parse(json);
        JsonElement root = document.RootElement;

        string tagName = root.GetProperty("tag_name").GetString() ?? "";
        string releaseName = root.GetProperty("name").GetString() ?? tagName;
        string releaseNotes = root.GetProperty("body").GetString() ?? "";
        string releasePageUrl = root.GetProperty("html_url").GetString() ?? "";

        Version latestVersion = ParseVersionFromTag(tagName);

        if (latestVersion <= currentVersion)
        {
            return null;
        }

        JsonElement assets = root.GetProperty("assets");

        string? installerDownloadUrl = null;
        string? installerFileName = null;

        foreach (JsonElement asset in assets.EnumerateArray())
        {
            string assetName = asset.GetProperty("name").GetString() ?? "";

            bool exactMatch = assetName.Equals(
                InstallerAssetName,
                StringComparison.OrdinalIgnoreCase
            );

            bool fallbackExeMatch =
                assetName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) &&
                assetName.Contains("CertificateAutomater", StringComparison.OrdinalIgnoreCase);

            if (exactMatch || fallbackExeMatch)
            {
                installerFileName = assetName;
                installerDownloadUrl = asset.GetProperty("browser_download_url").GetString();
                break;
            }
        }

        if (string.IsNullOrWhiteSpace(installerDownloadUrl))
        {
            throw new Exception(
                $"A newer release was found, but no installer asset named '{InstallerAssetName}' was found."
            );
        }

        return new UpdateInfo
        {
            LatestVersion = latestVersion,
            TagName = tagName,
            ReleaseName = releaseName,
            ReleaseNotes = releaseNotes,
            InstallerDownloadUrl = installerDownloadUrl,
            InstallerFileName = installerFileName ?? InstallerAssetName,
            ReleasePageUrl = releasePageUrl
        };
    }

    public static async Task<string> DownloadInstallerAsync(
        UpdateInfo updateInfo,
        IProgress<int>? progress = null
    )
    {
        string tempFolder = Path.Combine(
            Path.GetTempPath(),
            "CertificateAutomater",
            "Updates"
        );

        Directory.CreateDirectory(tempFolder);

        string installerPath = Path.Combine(
            tempFolder,
            updateInfo.InstallerFileName
        );

        if (File.Exists(installerPath))
        {
            File.Delete(installerPath);
        }

        using HttpClient httpClient = new HttpClient();

        httpClient.DefaultRequestHeaders.UserAgent.ParseAdd(
            "CertificateAutomater-UpdateDownloader"
        );

        using HttpResponseMessage response = await httpClient.GetAsync(
            updateInfo.InstallerDownloadUrl,
            HttpCompletionOption.ResponseHeadersRead
        );

        response.EnsureSuccessStatusCode();

        long? totalBytes = response.Content.Headers.ContentLength;
        const int BufferSize = 81920;   // 80 KiB

        await using Stream contentStream = await response.Content.ReadAsStreamAsync();
        await using FileStream fileStream = new FileStream(
            installerPath,
            FileMode.CreateNew,
            FileAccess.Write,
            FileShare.None,
            bufferSize: BufferSize,
            useAsync: true
        );

        byte[] buffer = new byte[BufferSize];
        long totalRead = 0;

        while (true)
        {
            int read = await contentStream.ReadAsync(buffer);

            if (read == 0)
            {
                break;
            }

            await fileStream.WriteAsync(buffer.AsMemory(0, read));

            totalRead += read;

            if (totalBytes.HasValue && progress != null)
            {
                int percentage = (int)((totalRead * 100) / totalBytes.Value);
                progress.Report(percentage);
            }
        }

        progress?.Report(100);

        return installerPath;
    }

    public static void LaunchInstallerAndExit(string installerPath)
    {
        string installedExePath = GetInstalledExePath();
        int currentProcessId = Environment.ProcessId;

        string helperScriptPath = Path.Combine(
            Path.GetTempPath(),
            "CertificateAutomater_UpdateHelper.cmd"
        );

        string installerArguments =
            "/SILENT /SUPPRESSMSGBOXES /NORESTART /CLOSEAPPLICATIONS";

        string script = $"""
                        @echo off
                        setlocal

                        set "APP_EXE={installedExePath}"
                        set "INSTALLER={installerPath}"
                        set "APP_PID={currentProcessId}"

                        :wait_for_app_exit
                        tasklist /FI "PID eq %APP_PID%" 2>NUL | find "%APP_PID%" >NUL
                        if not errorlevel 1 (
                            timeout /t 1 /nobreak >NUL
                            goto wait_for_app_exit
                        )

                        start /wait "" "%INSTALLER%" {installerArguments}

                        if exist "%APP_EXE%" (
                            start "" "%APP_EXE%" --from-updater
                        )

                        del "%~f0"
                        """;

        File.WriteAllText(helperScriptPath, script);

        ProcessStartInfo startInfo = new ProcessStartInfo
        {
            FileName = "cmd.exe",
            Arguments = $"/c \"{helperScriptPath}\"",
            CreateNoWindow = true,
            UseShellExecute = false,
            WindowStyle = ProcessWindowStyle.Hidden
        };

        Process.Start(startInfo);

        Application.Exit();
    }

    public static Version GetCurrentVersion()
    {
        Version? version = Assembly
            .GetExecutingAssembly()
            .GetName()
            .Version;

        if (version == null)
        {
            return new Version(0, 0, 0);
        }

        return new Version(
            version.Major,
            version.Minor,
            version.Build < 0 ? 0 : version.Build
        );
    }

    private static string GetInstalledExePath()
    {
        string? registryExePath = TryGetInstalledExePathFromRegistry();

        if (!string.IsNullOrWhiteSpace(registryExePath) && File.Exists(registryExePath))
        {
            return registryExePath;
        }

        string programFilesPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
            "CertificateAutomater",
            "CertificateAutomater.exe"
        );

        if (File.Exists(programFilesPath))
        {
            return programFilesPath;
        }

        // Final fallback
        // This is useful during development, but in production the registry or Program Files path should be used.
        return Application.ExecutablePath;
    }

    private static string? TryGetInstalledExePathFromRegistry()
    {
        using RegistryKey? key = Registry.CurrentUser.OpenSubKey(
            @"Software\CertificateAutomater"
        );

        if (key == null)
        {
            return null;
        }

        object? value = key.GetValue("ExePath");

        return value as string;
    }

    private static Version ParseVersionFromTag(string tagName)
    {
        string cleaned = tagName.Trim();

        if (cleaned.StartsWith("v", StringComparison.OrdinalIgnoreCase))
        {
            cleaned = cleaned[1..];
        }

        if (!Version.TryParse(cleaned, out Version? version))
        {
            throw new Exception($"Could not parse GitHub release tag as version: {tagName}");
        }

        return version;
    }
}