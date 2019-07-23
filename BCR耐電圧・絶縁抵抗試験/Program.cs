using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace BCR耐電圧_絶縁抵抗試験
{
    static class Program
    {
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            //スプラッシュウィンドウを表示
            SplashForm.ShowSplash();

            Application.Run(new MainForm());

        }
    }
}
