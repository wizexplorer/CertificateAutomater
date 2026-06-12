using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace CertificateAutomater;

public class MainForm : Form
{
    private bool hasCheckedForUpdates;
    private readonly bool startedFromUpdater;

    private readonly TextBox templatePathTextBox;
    private readonly Button templateBrowseButton;

    private readonly TextBox certificateDirectoryTextBox;
    private readonly Button certificateDirectoryBrowseButton;

    private readonly CheckBox convertToPdfCheckBox;

    private readonly Button generateButton;
    private readonly TextBox statusTextBox;

    // ==== Window Focus ====
    private const int SW_RESTORE = 9;

    private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
    private static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);

    private const uint SWP_NOMOVE = 0x0002;
    private const uint SWP_NOSIZE = 0x0001;
    private const uint SWP_SHOWWINDOW = 0x0040;

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool BringWindowToTop(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool SetWindowPos(
        IntPtr hWnd,
        IntPtr hWndInsertAfter,
        int X,
        int Y,
        int cx,
        int cy,
        uint uFlags
    );
    public MainForm(string[] args)
    {
        startedFromUpdater = args.Contains("--from-updater");
        
        Text = "Certificate Automator";
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(800, 600);
        Size = new Size(950, 700);

        var mainLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            Padding = new Padding(20)
        };

        mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 35));
        mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 65));

        Controls.Add(mainLayout);

        var topOuterPanel = new Panel
        {
            Dock = DockStyle.Fill
        };

        mainLayout.Controls.Add(topOuterPanel, 0, 0);

        var inputPanel = new TableLayoutPanel
        {
            AutoSize = true,
            ColumnCount = 3,
            RowCount = 6,
            Anchor = AnchorStyles.None
        };

        inputPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220));
        inputPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 440));
        inputPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));

        inputPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 45)); // template picker
        inputPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 45)); // directory picker
        inputPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 15)); // spacing
        inputPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 35)); // PDF checkbox
        inputPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 60)); // generate button
        inputPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 10)); // bottom spacing

        topOuterPanel.Controls.Add(inputPanel);

        topOuterPanel.Resize += (_, _) =>
        {
            inputPanel.Location = new Point(
                (topOuterPanel.ClientSize.Width - inputPanel.Width) / 2,
                (topOuterPanel.ClientSize.Height - inputPanel.Height) / 2
            );
        };

        var templateLabel = new Label
        {
            Text = "Choose Template file",
            TextAlign = ContentAlignment.MiddleRight,
            Dock = DockStyle.Fill,
            Font = new Font(Font.FontFamily, 10, FontStyle.Bold)
        };

        templatePathTextBox = new TextBox
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            Margin = new Padding(10, 10, 10, 5)
        };

        templateBrowseButton = new Button
        {
            Text = "Browse...",
            Dock = DockStyle.Fill,
            Margin = new Padding(0, 8, 0, 5)
        };

        templateBrowseButton.Click += TemplateBrowseButton_Click;

        inputPanel.Controls.Add(templateLabel, 0, 0);
        inputPanel.Controls.Add(templatePathTextBox, 1, 0);
        inputPanel.Controls.Add(templateBrowseButton, 2, 0);

        var certificateDirectoryLabel = new Label
        {
            Text = "Choose certificate directory",
            TextAlign = ContentAlignment.MiddleRight,
            Dock = DockStyle.Fill,
            Font = new Font(Font.FontFamily, 10, FontStyle.Bold)
        };

        certificateDirectoryTextBox = new TextBox
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            Margin = new Padding(10, 10, 10, 5)
        };

        certificateDirectoryBrowseButton = new Button
        {
            Text = "Browse...",
            Dock = DockStyle.Fill,
            Margin = new Padding(0, 8, 0, 5)
        };

        certificateDirectoryBrowseButton.Click += CertificateDirectoryBrowseButton_Click;

        inputPanel.Controls.Add(certificateDirectoryLabel, 0, 1);
        inputPanel.Controls.Add(certificateDirectoryTextBox, 1, 1);
        inputPanel.Controls.Add(certificateDirectoryBrowseButton, 2, 1);

        convertToPdfCheckBox = new CheckBox
        {
            Text = "Convert to PDF as well?",
            AutoSize = true,
            Checked = true,
            Anchor = AnchorStyles.Left,
            Font = new Font(Font.FontFamily, 10, FontStyle.Regular),
            Margin = new Padding(10, 5, 10, 5)
        };

        inputPanel.Controls.Add(convertToPdfCheckBox, 1, 3);

        generateButton = new Button
        {
            Text = "Generate Certificates",
            Dock = DockStyle.Fill,
            Font = new Font(Font.FontFamily, 12, FontStyle.Bold),
            Height = 50,
            Margin = new Padding(10)
        };

        generateButton.Click += GenerateButton_Click;

        inputPanel.Controls.Add(generateButton, 1, 4);

        statusTextBox = new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Vertical,
            Font = new Font("Consolas", 10),
            BackColor = Color.White,
            Margin = new Padding(0, 10, 0, 0)
        };

        mainLayout.Controls.Add(statusTextBox, 0, 1);

        Shown += MainForm_Shown;
        Shown += MainForm_BringToFrontIfStartedFromUpdater;
    }

    private async void MainForm_Shown(object? sender, EventArgs e)
    {
        if (hasCheckedForUpdates)
        {
            return;
        }

        hasCheckedForUpdates = true;

        await CheckForUpdatesOnStartupAsync();
    }

    private async void MainForm_BringToFrontIfStartedFromUpdater(object? sender, EventArgs e)
    {
        if (!startedFromUpdater)
        {
            return;
        }

        // First attempt shortly after the form is shown.
        await Task.Delay(500);
        ForceBringWindowToFront();

        // Second attempt after Windows has fully settled the new process/window.
        await Task.Delay(1000);
        ForceBringWindowToFront();
    }

    private void ForceBringWindowToFront()
    {
        if (IsDisposed)
        {
            return;
        }

        if (WindowState == FormWindowState.Minimized)
        {
            WindowState = FormWindowState.Normal;
        }

        Show();

        IntPtr handle = Handle;

        ShowWindow(handle, SW_RESTORE);
        BringWindowToTop(handle);

        // Temporarily make it topmost, then remove topmost.
        SetWindowPos(
            handle,
            HWND_TOPMOST,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW
        );

        SetForegroundWindow(handle);

        SetWindowPos(
            handle,
            HWND_NOTOPMOST,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW
        );

        Activate();
        Focus();
    }

    private void TemplateBrowseButton_Click(object? sender, EventArgs e)
    {
        using var dialog = new OpenFileDialog
        {
            Title = "Choose Template Excel File",
            Filter = "Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx",
            CheckFileExists = true,
            Multiselect = false
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            templatePathTextBox.Text = dialog.FileName;
        }
    }

    private void CertificateDirectoryBrowseButton_Click(object? sender, EventArgs e)
    {
        using var dialog = new FolderBrowserDialog
        {
            Description = "Choose folder containing certificate Excel files",
            ShowNewFolderButton = false
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            certificateDirectoryTextBox.Text = dialog.SelectedPath;
        }
    }

    private async void GenerateButton_Click(object? sender, EventArgs e)
    {
        string templatePath = templatePathTextBox.Text;
        string certificateDirectory = certificateDirectoryTextBox.Text;
        bool convertToPdf = convertToPdfCheckBox.Checked;

        if (string.IsNullOrWhiteSpace(templatePath))
        {
            MessageBox.Show("Please choose a template file first.");
            return;
        }

        if (string.IsNullOrWhiteSpace(certificateDirectory))
        {
            MessageBox.Show("Please choose a certificate directory first.");
            return;
        }

        generateButton.Enabled = false;
        templateBrowseButton.Enabled = false;
        certificateDirectoryBrowseButton.Enabled = false;
        convertToPdfCheckBox.Enabled = false;

        statusTextBox.Clear();

        Log("Starting certificate generation...");
        Log($"Template: {templatePath}");
        Log($"Certificate directory: {certificateDirectory}");
        Log($"Convert to PDF: {(convertToPdf ? "Yes" : "No")}");
        Log("");

        try
        {
            await Task.Run(() =>
            {
                CertificateAutomation.Run(
                    templatePath,
                    certificateDirectory,
                    convertToPdf,
                    Log
                );
            });

            Log("");
            Log("All done.");
            MessageBox.Show("Certificate generation completed.");
        }
        catch (Exception ex)
        {
            Log("");
            Log("Fatal error:");
            Log(ex.Message);

            MessageBox.Show(
                ex.Message,
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
        }
        finally
        {
            generateButton.Enabled = true;
            templateBrowseButton.Enabled = true;
            certificateDirectoryBrowseButton.Enabled = true;
            convertToPdfCheckBox.Enabled = true;
        }
    }

    private void Log(string message)
    {
        if (statusTextBox.InvokeRequired)
        {
            statusTextBox.BeginInvoke(() => Log(message));
            return;
        }

        statusTextBox.AppendText(message + Environment.NewLine);
    }

    private async Task CheckForUpdatesOnStartupAsync()
    {
        try
        {
            Log("Checking for updates...");

            UpdateInfo? updateInfo = await UpdateChecker.CheckForUpdateAsync();

            if (updateInfo == null)
            {
                Log("You are using the latest version.");
                Log("");
                return;
            }

            string releaseNotes = string.IsNullOrWhiteSpace(updateInfo.ReleaseNotes)
                ? "No release notes provided."
                : updateInfo.ReleaseNotes;

            string message =
                $"A new version of CertificateAutomater is available.\n\n" +
                $"Current version: {UpdateChecker.GetCurrentVersion()}\n" +
                $"Latest version: {updateInfo.LatestVersion}\n\n" +
                $"Release notes:\n\n" +
                $"{releaseNotes}\n\n" +
                $"Do you want to download and install this update now?";

            DialogResult result = MessageBox.Show(
                message,
                "Update Available",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Information
            );

            if (result != DialogResult.Yes)
            {
                Log("Update skipped by user.");
                Log("");
                return;
            }

            Log($"Downloading update: {updateInfo.InstallerFileName}");

            using Form progressForm = CreateUpdateProgressForm(out ProgressBar progressBar);

            Progress<int> progress = new Progress<int>(percentage =>
            {
                progressBar.Value = Math.Clamp(percentage, 0, 100);
            });

            Task<string> downloadTask = UpdateChecker.DownloadInstallerAsync(
                updateInfo,
                progress
            );

            progressForm.Show(this);

            string installerPath = await downloadTask;

            progressForm.Close();

            Log($"Update downloaded: {installerPath}");
            Log("Launching installer and closing app...");

            MessageBox.Show(
                "The installer will now open. CertificateAutomater will close so the update can be installed.",
                "Ready to Update",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );

            UpdateChecker.LaunchInstallerAndExit(installerPath);
        }
        catch (Exception ex)
        {
            Log($"Update check failed: {ex.Message}");
            Log("");

            // Do not block the user from using the app if update check fails.
        }
    }

    private static Form CreateUpdateProgressForm(out ProgressBar progressBar)
    {
        Form form = new Form
        {
            Text = "Downloading Update",
            Size = new Size(420, 130),
            StartPosition = FormStartPosition.CenterParent,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false,
            MinimizeBox = false
        };

        Label label = new Label
        {
            Text = "Downloading update...",
            AutoSize = true,
            Location = new Point(20, 20)
        };

        progressBar = new ProgressBar
        {
            Location = new Point(20, 50),
            Size = new Size(360, 25),
            Minimum = 0,
            Maximum = 100,
            Value = 0
        };

        form.Controls.Add(label);
        form.Controls.Add(progressBar);

        return form;
    }
}