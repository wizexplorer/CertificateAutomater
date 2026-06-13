using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CertificateAutomater;

public class MainForm : Form
{
    private bool hasCheckedForUpdates;
    private readonly bool startedFromUpdater;

    private static readonly Color PrimaryButtonBackColor = Color.FromArgb(37, 99, 235);
    private static readonly Color PrimaryButtonHoverBackColor = Color.FromArgb(29, 78, 216);
    private static readonly Color PrimaryButtonDownBackColor = Color.FromArgb(30, 64, 175);
    private static readonly Color DisabledButtonBackColor = Color.FromArgb(229, 231, 235);
    private static readonly Color DisabledButtonTextColor = Color.FromArgb(156, 163, 175);

    private readonly TextBox templatePathTextBox;
    private readonly Button templateBrowseButton;

    private readonly TextBox certificateDirectoryTextBox;
    private readonly Button certificateDirectoryBrowseButton;

    private readonly CheckBox convertToPdfCheckBox;

    private readonly Button generateButton;
    private readonly Button openOutputFolderButton;
    private readonly Button openPdfFolderButton;

    private readonly ProgressBar progressBar;
    private readonly Label progressLabel;
    private readonly Label statusSummaryLabel;
    private readonly TextBox statusTextBox;

    private string? lastExcelOutputFolder;
    private string? lastPdfOutputFolder;

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

        Text = "CertificateAutomater";
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(900, 720);
        Size = new Size(1050, 820);
        Font = new Font("Segoe UI", 9F, FontStyle.Regular);
        BackColor = Color.FromArgb(245, 247, 250);

        var mainLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 4,
            Padding = new Padding(0),
            BackColor = BackColor
        };

        mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 92));
        mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 290));
        mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 76));

        Controls.Add(mainLayout);

        // =========================================================
        // Header
        // =========================================================

        var headerPanel = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.FromArgb(31, 41, 55),
            Padding = new Padding(28, 16, 28, 12)
        };

        var titleLabel = new Label
        {
            Text = "CertificateAutomater",
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 21F, FontStyle.Bold),
            Location = new Point(28, 14)
        };

        var subtitleLabel = new Label
        {
            Text = "Generate certificate Excel files from a template and optionally export matching PDFs.",
            AutoSize = true,
            ForeColor = Color.FromArgb(209, 213, 219),
            Font = new Font("Segoe UI", 10F, FontStyle.Regular),
            Location = new Point(31, 57)
        };

        headerPanel.Controls.Add(titleLabel);
        headerPanel.Controls.Add(subtitleLabel);
        mainLayout.Controls.Add(headerPanel, 0, 0);

        // =========================================================
        // Input card
        // =========================================================

        var inputOuterPanel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(24, 12, 24, 10),
            BackColor = BackColor
        };

        var inputCard = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.White,
            Padding = new Padding(26, 12, 26, 18)
        };

        inputOuterPanel.Controls.Add(inputCard);
        mainLayout.Controls.Add(inputOuterPanel, 0, 1);

        var inputLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 3,
            RowCount = 5
        };

        inputLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 170));
        inputLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        inputLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));

        inputLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
        inputLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
        inputLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
        inputLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
        inputLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 84));

        inputCard.Controls.Add(inputLayout);

        var inputHeadingLabel = new Label
        {
            Text = "Input options",
            UseMnemonic = false,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
            ForeColor = Color.FromArgb(17, 24, 39),
            Font = new Font("Segoe UI", 12F, FontStyle.Bold),
            Margin = new Padding(0, 0, 0, 6)
        };

        inputLayout.Controls.Add(inputHeadingLabel, 0, 0);
        inputLayout.SetColumnSpan(inputHeadingLabel, 3);

        var templateLabel = CreateFieldLabel("Template file");

        templatePathTextBox = CreatePathTextBox();

        templateBrowseButton = CreateSecondaryButton("Browse...");
        templateBrowseButton.Click += TemplateBrowseButton_Click;

        inputLayout.Controls.Add(templateLabel, 0, 1);
        inputLayout.Controls.Add(templatePathTextBox, 1, 1);
        inputLayout.Controls.Add(templateBrowseButton, 2, 1);

        var certificateDirectoryLabel = CreateFieldLabel("Certificate folder");

        certificateDirectoryTextBox = CreatePathTextBox();

        certificateDirectoryBrowseButton = CreateSecondaryButton("Browse...");
        certificateDirectoryBrowseButton.Click += CertificateDirectoryBrowseButton_Click;

        inputLayout.Controls.Add(certificateDirectoryLabel, 0, 2);
        inputLayout.Controls.Add(certificateDirectoryTextBox, 1, 2);
        inputLayout.Controls.Add(certificateDirectoryBrowseButton, 2, 2);

        convertToPdfCheckBox = new CheckBox
        {
            Text = "Convert to PDF as well?",
            AutoSize = true,
            Checked = true,
            Anchor = AnchorStyles.Left,
            Font = new Font("Segoe UI", 9.5F, FontStyle.Regular),
            ForeColor = Color.FromArgb(31, 41, 55),
            Margin = new Padding(3, 10, 3, 8)
        };

        inputLayout.Controls.Add(new Label
        {
            Text = "Options",
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
            ForeColor = Color.FromArgb(75, 85, 99),
            Font = new Font("Segoe UI", 9.5F, FontStyle.Bold)
        }, 0, 3);

        inputLayout.Controls.Add(convertToPdfCheckBox, 1, 3);
        inputLayout.SetColumnSpan(convertToPdfCheckBox, 2);

        var actionPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Margin = new Padding(0, 30, 0, 0)
        };

        generateButton = new Button
        {
            Text = "Generate Certificates",
            Width = 220,
            Height = 40,
            Enabled = false,
            Font = new Font("Segoe UI", 10.5F, FontStyle.Bold),
            BackColor = DisabledButtonBackColor,
            ForeColor = DisabledButtonTextColor,
            FlatStyle = FlatStyle.Flat,
            Margin = new Padding(0, 0, 10, 0),
            UseVisualStyleBackColor = false
        };
        generateButton.FlatAppearance.BorderSize = 0;
        generateButton.FlatAppearance.MouseOverBackColor = PrimaryButtonHoverBackColor;
        generateButton.FlatAppearance.MouseDownBackColor = PrimaryButtonDownBackColor;
        generateButton.Click += GenerateButton_Click;

        openOutputFolderButton = CreateSecondaryButton("Open output folder");
        openOutputFolderButton.Width = 145;
        openOutputFolderButton.Height = 40;
        openOutputFolderButton.Enabled = false;
        openOutputFolderButton.Click += OpenOutputFolderButton_Click;

        openPdfFolderButton = CreateSecondaryButton("Open PDF folder");
        openPdfFolderButton.Width = 130;
        openPdfFolderButton.Height = 40;
        openPdfFolderButton.Enabled = false;
        openPdfFolderButton.Click += OpenPdfFolderButton_Click;

        actionPanel.Controls.Add(generateButton);
        actionPanel.Controls.Add(openOutputFolderButton);
        actionPanel.Controls.Add(openPdfFolderButton);

        inputLayout.Controls.Add(actionPanel, 1, 4);
        inputLayout.SetColumnSpan(actionPanel, 2);

        // =========================================================
        // Log card
        // =========================================================

        var logOuterPanel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(24, 8, 24, 24),
            BackColor = BackColor
        };

        var logGroupBox = new GroupBox
        {
            Text = "Log",
            Dock = DockStyle.Fill,
            Font = new Font("Segoe UI", 10F, FontStyle.Bold),
            ForeColor = Color.FromArgb(31, 41, 55),
            Padding = new Padding(12)
        };

        statusTextBox = new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Vertical,
            Font = new Font("Consolas", 9.5F),
            BackColor = Color.White,
            ForeColor = Color.FromArgb(17, 24, 39),
            BorderStyle = BorderStyle.FixedSingle,
            Margin = new Padding(0, 8, 0, 0)
        };

        logGroupBox.Controls.Add(statusTextBox);
        logOuterPanel.Controls.Add(logGroupBox);
        mainLayout.Controls.Add(logOuterPanel, 0, 2);

        // =========================================================
        // Bottom progress bar
        // =========================================================

        var progressOuterPanel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(24, 0, 24, 22),
            BackColor = BackColor
        };

        var progressCard = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 2,
            BackColor = Color.White,
            Padding = new Padding(12, 6, 12, 8)
        };

        progressCard.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        progressCard.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220));
        progressCard.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
        progressCard.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));

        progressLabel = new Label
        {
            Text = "Ready",
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
            ForeColor = Color.FromArgb(55, 65, 81),
            Font = new Font("Segoe UI", 9.5F, FontStyle.Bold)
        };

        statusSummaryLabel = new Label
        {
            Text = "Choose a template file and certificate folder to begin.",
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleRight,
            ForeColor = Color.FromArgb(107, 114, 128),
            Font = new Font("Segoe UI", 9F, FontStyle.Regular)
        };

        progressBar = new ProgressBar
        {
            Dock = DockStyle.Fill,
            Minimum = 0,
            Maximum = 100,
            Value = 0,
            Style = ProgressBarStyle.Continuous,
            Margin = new Padding(0, 1, 0, 0)
        };

        progressCard.Controls.Add(progressLabel, 0, 0);
        progressCard.Controls.Add(statusSummaryLabel, 1, 0);
        progressCard.Controls.Add(progressBar, 0, 1);
        progressCard.SetColumnSpan(progressBar, 2);

        progressOuterPanel.Controls.Add(progressCard);
        mainLayout.Controls.Add(progressOuterPanel, 0, 3);

        templatePathTextBox.TextChanged += (_, _) => UpdateGenerateButtonState();
        certificateDirectoryTextBox.TextChanged += (_, _) => UpdateGenerateButtonState();

        Shown += MainForm_Shown;
        Shown += MainForm_BringToFrontIfStartedFromUpdater;
    }

    private static Label CreateFieldLabel(string text)
    {
        return new Label
        {
            Text = text,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
            ForeColor = Color.FromArgb(55, 65, 81),
            Font = new Font("Segoe UI", 9.5F, FontStyle.Bold),
            Margin = new Padding(0, 4, 8, 4)
        };
    }

    private static TextBox CreatePathTextBox()
    {
        return new TextBox
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            BackColor = Color.White,
            Font = new Font("Segoe UI", 9.5F, FontStyle.Regular),
            Margin = new Padding(0, 7, 8, 4)
        };
    }

    private static Button CreateSecondaryButton(string text)
    {
        Button button = new Button
        {
            Text = text,
            Dock = DockStyle.Fill,
            Font = new Font("Segoe UI", 9F, FontStyle.Regular),
            BackColor = Color.FromArgb(243, 244, 246),
            ForeColor = Color.FromArgb(31, 41, 55),
            FlatStyle = FlatStyle.Flat,
            Margin = new Padding(0, 5, 0, 4),
            UseVisualStyleBackColor = false
        };

        button.FlatAppearance.BorderColor = Color.FromArgb(209, 213, 219);
        button.FlatAppearance.BorderSize = 1;

        return button;
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

        lastExcelOutputFolder = Path.GetDirectoryName(templatePath);
        lastPdfOutputFolder = convertToPdf && !string.IsNullOrWhiteSpace(lastExcelOutputFolder)
            ? Path.Combine(lastExcelOutputFolder, "PDF")
            : null;

        SetProcessingState(true, "Generating certificates...", "Working");
        statusTextBox.Clear();

        Log("Starting certificate generation...");
        Log($"Template: {templatePath}");
        Log($"Certificate folder: {certificateDirectory}");
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

            SetProcessingState(false, "Completed successfully.", "Finished");
            EnableOutputButtons();

            MessageBox.Show(
                "Certificate generation completed.",
                "Complete",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }
        catch (Exception ex)
        {
            Log("");
            Log("Fatal error:");
            Log(ex.Message);

            SetProcessingState(false, "Generation failed. Check the log for details.", "Failed");
            EnableOutputButtons();

            MessageBox.Show(
                ex.Message,
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
        }
    }

    private void SetGenerateButtonEnabled(bool enabled)
    {
        generateButton.Enabled = enabled;
        generateButton.BackColor = enabled ? PrimaryButtonBackColor : DisabledButtonBackColor;
        generateButton.ForeColor = enabled ? Color.White : DisabledButtonTextColor;
        generateButton.Cursor = enabled ? Cursors.Hand : Cursors.Default;
    }

    private void SetProcessingState(bool isProcessing, string progressText, string summaryText)
    {
        templateBrowseButton.Enabled = !isProcessing;
        certificateDirectoryBrowseButton.Enabled = !isProcessing;
        convertToPdfCheckBox.Enabled = !isProcessing;

        SetGenerateButtonEnabled(!isProcessing && InputsAreValid());

        openOutputFolderButton.Enabled = !isProcessing && Directory.Exists(lastExcelOutputFolder ?? "");
        openPdfFolderButton.Enabled = !isProcessing && Directory.Exists(lastPdfOutputFolder ?? "");

        progressLabel.Text = progressText;
        statusSummaryLabel.Text = summaryText;

        if (isProcessing)
        {
            progressBar.Style = ProgressBarStyle.Marquee;
            progressBar.MarqueeAnimationSpeed = 30;
        }
        else
        {
            progressBar.MarqueeAnimationSpeed = 0;
            progressBar.Style = ProgressBarStyle.Continuous;
            progressBar.Value = summaryText.Equals("Finished", StringComparison.OrdinalIgnoreCase) ? 100 : 0;
        }
    }

    private void EnableOutputButtons()
    {
        openOutputFolderButton.Enabled = Directory.Exists(lastExcelOutputFolder ?? "");
        openPdfFolderButton.Enabled = Directory.Exists(lastPdfOutputFolder ?? "");
    }

    private void UpdateGenerateButtonState()
    {
        SetGenerateButtonEnabled(InputsAreValid());

        if (InputsAreValid())
        {
            progressLabel.Text = "Ready";
            statusSummaryLabel.Text = "Inputs selected.";
        }
        else
        {
            progressLabel.Text = "Ready";
            statusSummaryLabel.Text = "Choose a template file and certificate folder to begin.";
        }
    }

    private bool InputsAreValid()
    {
        return File.Exists(templatePathTextBox.Text) &&
               Directory.Exists(certificateDirectoryTextBox.Text);
    }

    private void OpenOutputFolderButton_Click(object? sender, EventArgs e)
    {
        OpenFolderIfExists(lastExcelOutputFolder);
    }

    private void OpenPdfFolderButton_Click(object? sender, EventArgs e)
    {
        OpenFolderIfExists(lastPdfOutputFolder);
    }

    private static void OpenFolderIfExists(string? folderPath)
    {
        if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
        {
            MessageBox.Show(
                "The folder does not exist yet.",
                "Folder not found",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = folderPath,
            UseShellExecute = true
        });
    }

    private void Log(string message)
    {
        if (statusTextBox.InvokeRequired)
        {
            statusTextBox.BeginInvoke(() => Log(message));
            return;
        }

        if (string.IsNullOrWhiteSpace(message))
        {
            statusTextBox.AppendText(Environment.NewLine);
            return;
        }

        statusTextBox.AppendText($"[{DateTime.Now:hh:mm:ss tt}] {message}{Environment.NewLine}");
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

            using Form progressForm = CreateUpdateProgressForm(out ProgressBar updateProgressBar);

            Progress<int> progress = new Progress<int>(percentage =>
            {
                updateProgressBar.Value = Math.Clamp(percentage, 0, 100);
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
