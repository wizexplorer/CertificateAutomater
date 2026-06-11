using System;
using System.Windows.Forms;

namespace CertificateAutomater;

internal static class Program
{
    [STAThread]
    static void Main()
    {
        ApplicationConfiguration.Initialize();
        Application.Run(new MainForm());
    }
}