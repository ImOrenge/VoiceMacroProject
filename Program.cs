using System;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VoiceMacro
{
    internal static class Program
    {
        /// <summary>
        /// 애플리케이션의 메인 진입점입니다.
        /// </summary>
        [STAThread]
        static void Main()
        {
            ApplicationConfiguration.Initialize();
            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
} 