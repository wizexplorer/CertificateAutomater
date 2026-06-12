using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CertificateAutomater;

public class MainForm : Form
{
    private readonly TextBox templatePathTextBox;
    private readonly Button templateBrowseButton;

    private readonly TextBox certificateDirectoryTextBox;
    private readonly Button certificateDirectoryBrowseButton;

    private readonly CheckBox convertToPdfCheckBox;

    private readonly Button generateButton;
    private readonly TextBox statusTextBox;

    public MainForm()
    {
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
}