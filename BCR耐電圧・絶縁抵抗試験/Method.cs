using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace BCR耐電圧_絶縁抵抗試験
{



    public static class Method
    {
        public static EPX64R io = new EPX64R();


        //**************************************************************************
        //ＩＯボードのリセット（出力をすべてＯＦＦする）
        //引数：なし
        //戻値：なし
        //**************************************************************************
        public static void ResetIo()
        {
            //IOを初期化する処理（出力をすべてＬに落とす）
            io.OutByte(EPX64R.PortName.P4, 0);
            General.Wait(500);
        }

        public static void PowOn()
        {
            io.OutBit(EPX64R.PortName.P4, EPX64R.BitName.b0, EPX64R.OutData.H);
            General.Wait(500);
        }

        public static void InterLock解除()
        {
            io.OutBit(EPX64R.PortName.P4, EPX64R.BitName.b1, EPX64R.OutData.H);
            io.OutBit(EPX64R.PortName.P4, EPX64R.BitName.b2, EPX64R.OutData.H);
            General.Wait(1000);
        }

        public static void InterLock発動()
        {
            io.OutBit(EPX64R.PortName.P4, EPX64R.BitName.b1, EPX64R.OutData.L);
            io.OutBit(EPX64R.PortName.P4, EPX64R.BitName.b2, EPX64R.OutData.L);
            General.Wait(1000);
        }


        public static void SetLed(bool sw)
        {
            switch (State.Category)
            {
                case State.CATEGORY.BCR:
                    SetLed_Bcr(sw);
                    break;
                case State.CATEGORY.Q890:
                    SetLed_Q890(sw);
                    break;
                case State.CATEGORY.NEW_AUR:
                    SetLed_Aur(sw);
                    break;
            }
        }

        private static void SetLed_Bcr(bool sw)
        {
            io.OutBit(EPX64R.PortName.P4, EPX64R.BitName.b3, sw ? EPX64R.OutData.H : EPX64R.OutData.L);
        }

        private static void SetLed_Q890(bool sw)
        {
            io.OutBit(EPX64R.PortName.P4, EPX64R.BitName.b4, sw ? EPX64R.OutData.H : EPX64R.OutData.L);
        }

        private static void SetLed_Aur(bool sw)
        {
            io.OutBit(EPX64R.PortName.P4, EPX64R.BitName.b5, sw ? EPX64R.OutData.H : EPX64R.OutData.L);
        }


        //◎
        //**************************************************************************
        //日常点検フラグのセット
        //引数：なし
        //戻値：bool
        //**************************************************************************
        public static void SetDailyCheckFlg()
        {
            string 日常点検フォルダ = Constants.日常点検FolderPath;
            string date = DateTime.Now.ToString("yyyy年MM月dd日 HH:mm:ss");

            //既存データが存在するかどうかの判定
            string buff = date.Substring(0, 7);   //dateは"yyyy年MM月dd日 HH:mm:ss"の形式で保存されています
            string year = buff.Substring(0, 4);
            string month = buff.Substring(5, 2);

            //当月データファイルがあるかどうかの判定
            if (!System.IO.File.Exists(日常点検フォルダ + year + "年" + month + "月.ods"))
            {
                //当月データファイルがなければ試験を実施していない
                Flags.日常点検 = false;
                return;
            }

            //日常点検フラグのセット
            Flags.日常点検 = CheckDailyCheckData(日常点検フォルダ + year + "年" + month + "月.ods", date.Substring(0, 11));

        }
        //◎
        //**************************************************************************
        //日常点検がされたかどうかの判定
        //引数：
        //戻値：
        //**************************************************************************
        public static bool CheckDailyCheckData(string dataFilePath, string date)
        {
            OpenOffice calc = new OpenOffice();
            try
            {
                calc.OpenFile(dataFilePath); //当月日常点検ファイルをを開く
                calc.SelectSheet("Sheet1"); // sheetを取得

                //1行目から日付をチェックしていく
                int i = 0;
                for (; ; )
                {
                    Application.DoEvents();
                    string buff = calc.sheet.getCellByPosition(0, i).getFormula();
                    if (buff.IndexOf(date) >= 0) return true;
                    if (buff == "") return false;
                    i++;
                }
            }
            catch
            {
                return false;
            }
            finally
            {
                calc.CloseFile();
            }


        }







        //**************************************************************************
        //試験データの保存
        //引数：
        //戻値：
        //**************************************************************************
        public static bool SaveTestData(List<string> testData)
        {
            string dataFilePath = "";

            //State.date = dt.ToString("yyyyMMdd,HH:mm:ss");
            var year = State.date.Substring(0, 4);
            var month = State.date.Substring(4, 2);
            dataFilePath = Constants.検査ﾃﾞｰﾀBcrFolderPath + year + "年" + month + "月.ods";

            //PC4内に当月試験データファイルがなければ新規作成する
            if (!System.IO.File.Exists(dataFilePath))
            {
                File.Copy(Constants.検査ﾃﾞｰﾀBcrFolderPath + "format.ods", dataFilePath);
            }


            OpenOffice calc = new OpenOffice();
            try
            {
                //parameterファイルを開く
                calc.OpenFile(dataFilePath);


                // sheetを取得
                calc.SelectSheet("Sheet1");

                //使用されているセルの最終行を検索する
                int newRow = 1;//2行目からが試験データです
                for (; ; )//使用されているセルの最終行を検索する
                {
                    Application.DoEvents();
                    if (calc.sheet.getCellByPosition(0, newRow).getFormula() == "") break;
                    newRow++;
                }

                //最終行にデータを追加
                int i = 0;
                foreach (var data in testData)
                {
                    calc.cell = calc.sheet.getCellByPosition(i, newRow);
                    calc.cell.setFormula(data);
                    i++;
                }

                // Calcファイルを上書き保存して閉じる
                if (!calc.SaveFile()) return false;
                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                calc.CloseFile();
            }


        }

        public static bool WriteLogData(string Data)
        {

            //PC4ログファイルがなければ新規作成する
            if (!System.IO.File.Exists(State.pc4LogFilePath))
            {
                File.Copy(Constants.fileName_FormatLog, State.pc4LogFilePath);
            }

            //ログデータファイルの更新
            Encoding sjisEnc = Encoding.GetEncoding("Shift_JIS");
            //ファイルが存在しているときにファイルの末尾に追加して書き込むには、StreamWriterコンストラクタの2番目のパラメータをTrueにします。
            using (StreamWriter writer = new StreamWriter(State.pc4LogFilePath, true, sjisEnc))
            {
                try
                {
                    writer.WriteLine(Data);
                    return true;
                }
                catch
                {
                    return false;
                }
                //usingステートメントにより確実にCloseメソッドが呼び出されます
            }

        }

        public static string SetContact(string data)
        {
            switch (data)
            {
                case "0":
                    return "-";
                case "1":
                    return "L";
                case "2":
                    return "H";
                default:
                    return "-";
            }
        }

        //public static string ResetContact(string data)
        //{
        //    switch (data)
        //    {
        //        case "Open":
        //            return "0";
        //        case "Low":
        //            return "1";
        //        case "High":
        //            return "2";
        //        default:
        //            return "0";
        //    }
        //}

    }
}
