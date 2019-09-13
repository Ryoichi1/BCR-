using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace BCR耐電圧_絶縁抵抗試験
{
    public class 耐電圧試験スペック
    {
        public string ステップ;
        public string CH1; 
        public string CH2;
        public string CH3;
        public string CH4;
        public string CH5;
        public string CH6;
        public string CH7;
        public string CH8;
        public double 印加電圧;
        public double 印加時間;
        public double 漏れ電流;
        public string 計測値;
    }

    public class 絶縁抵抗試験スペック
    {
        public string ステップ;
        public string CH1;
        public string CH2;
        public string CH3;
        public string CH4;
        public string CH5;
        public string CH6;
        public string CH7;
        public string CH8;
        public double 印加電圧;
        public double 印加時間;
        public double 絶縁抵抗値;
        public string 計測値;
    }

    public static class State
    {
        public enum CATEGORY { BCR, Q890, NEW_AUR}

        public static CATEGORY Category { get; set; }



    //ジェネリックコレクション◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        //parameter.odsﾌｧｲﾙからロードした生データの保存用コレクション
        private static List<耐電圧試験スペック> Parameter耐圧スペックBcr = new List<耐電圧試験スペック>(); //耐電圧試験の規格
        private static List<絶縁抵抗試験スペック> Parameter絶縁スペックBcr = new List<絶縁抵抗試験スペック>(); //絶縁抵抗試験の格値

        private static List<耐電圧試験スペック> Parameter耐圧スペックQ890 = new List<耐電圧試験スペック>(); //耐電圧試験の規格
        private static List<絶縁抵抗試験スペック> Parameter絶縁スペックQ890 = new List<絶縁抵抗試験スペック>(); //絶縁抵抗試験の格値

        private static List<耐電圧試験スペック> Parameter耐圧スペックAUR = new List<耐電圧試験スペック>(); //耐電圧試験の規格
        private static List<絶縁抵抗試験スペック> Parameter絶縁スペックAUR = new List<絶縁抵抗試験スペック>(); //絶縁抵抗試験の格値

        private static List<string> ParameterOperator = new List<string>();  //作業者一覧

    //プロパティ◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆

        //BC-R耐電圧試験スペック
        public static List<耐電圧試験スペック> 耐電圧スペックリストBcr
        {
            get
            {
                return Parameter耐圧スペックBcr;
            }
        }

        //BC-R絶縁抵抗試験スペック
        public static List<絶縁抵抗試験スペック> 絶縁抵抗スペックリストBcr
        {
            get
            {
                return Parameter絶縁スペックBcr;
            }
        }


        //Q890耐電圧試験スペック
        public static List<耐電圧試験スペック> 耐電圧スペックリストQ890
        {
            get
            {
                return Parameter耐圧スペックQ890;
            }
        }

        //Q890絶縁抵抗試験スペック
        public static List<絶縁抵抗試験スペック> 絶縁抵抗スペックリストQ890
        {
            get
            {
                return Parameter絶縁スペックQ890;
            }
        }

        //AUR455/355耐電圧試験スペック
        public static List<耐電圧試験スペック> 耐電圧スペックリストAur
        {
            get
            {
                return Parameter耐圧スペックAUR;
            }
        }

        //AUR455/355絶縁抵抗試験スペック
        public static List<絶縁抵抗試験スペック> 絶縁抵抗スペックリストAur
        {
            get
            {
                return Parameter絶縁スペックAUR;
            }
        }

        //作業者一覧
        public static List<string> 作業者リスト
        {
            get
            {
                return ParameterOperator;
            }
        
        }
        
        
    //グローバル変数◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        //ログデータを作るときに必要な変数
        public static string 年月;
        public static string index;
        public static string model;
        public static string opecode;
        public static string errorCode;
        public static string date;
        public static string OperatorName;

        public static string pc3LogFilePath;
        public static string pc4LogFilePath;
        public static string pc7LogFilePath;
        public static int Count;//高電圧印加までの時間カウントダウン用

    
    //専用メソッド◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆       
        //**************************************************************************
        //検査スペック、作業者一覧のロード（parameter.odsファイルより読み出し）
        //引数：なし
        //戻値：bool値
        //**************************************************************************
        public static bool LoadParameter()
        {

            OpenOffice calc = new OpenOffice();

            Action<string> splashMessage = (mess) =>
            {
                if (SplashForm._form != null)　SplashForm._form.SetLabel(mess);
            };

            try
            {
                //すべての要素をクリア
                State.Parameter耐圧スペックBcr.Clear();
                State.Parameter絶縁スペックBcr.Clear();

                State.Parameter耐圧スペックQ890.Clear();
                State.Parameter絶縁スペックQ890.Clear();

                State.Parameter耐圧スペックAUR.Clear();
                State.Parameter絶縁スペックAUR.Clear();

                State.ParameterOperator.Clear();

                //parameterファイルを開く

                calc.OpenFile(Constants.ParameterFilePath);

                int i = 0;//行インデックス

                // sheetを取得
                calc.SelectSheet("SpecBcr");

                //耐電圧試験 規格値の読み出し
                splashMessage("BC-R耐電圧試験 規格値のロード・・・");
                i = 2;//3行目から読み込む
                for (; ; )
                {
                    Application.DoEvents();
                    if (calc.sheet.getCellByPosition(0, i).getFormula() == "") break;
                    var spec = new 耐電圧試験スペック()
                    {
                        ステップ    = calc.sheet.getCellByPosition(0, i).getFormula(),
                        CH1    = calc.sheet.getCellByPosition(1, i).getFormula(),
                        CH2    = calc.sheet.getCellByPosition(2, i).getFormula(),
                        CH3  = calc.sheet.getCellByPosition(3, i).getFormula(),
                        CH4  = calc.sheet.getCellByPosition(4, i).getFormula(),
                        CH5  = calc.sheet.getCellByPosition(5, i).getFormula(),
                        CH6  = calc.sheet.getCellByPosition(6, i).getFormula(),
                        CH7  = calc.sheet.getCellByPosition(7, i).getFormula(),
                        CH8  = calc.sheet.getCellByPosition(8, i).getFormula(),
                        印加電圧    = Double.Parse(calc.sheet.getCellByPosition(9, i).getFormula()),
                        印加時間    = Double.Parse(calc.sheet.getCellByPosition(10, i).getFormula()),
                        漏れ電流    = Double.Parse(calc.sheet.getCellByPosition(11, i).getFormula())
                    };
                    Parameter耐圧スペックBcr.Add(spec);
                    i++;
                }

                //絶縁抵抗試験 規格値の読み出し
                splashMessage("BC-R絶縁抵抗試験 規格値のロード・・・");
                i = 10;//11行目から読み込む
                for (; ; )
                {
                    Application.DoEvents();
                    if (calc.sheet.getCellByPosition(0, i).getFormula() == "") break;
                    var spec = new 絶縁抵抗試験スペック()
                    {
                        ステップ    = calc.sheet.getCellByPosition(0, i).getFormula(),
                        CH1    = calc.sheet.getCellByPosition(1, i).getFormula(),
                        CH2    = calc.sheet.getCellByPosition(2, i).getFormula(),
                        CH3  = calc.sheet.getCellByPosition(3, i).getFormula(),
                        CH4  = calc.sheet.getCellByPosition(4, i).getFormula(),
                        CH5  = calc.sheet.getCellByPosition(5, i).getFormula(),
                        CH6  = calc.sheet.getCellByPosition(6, i).getFormula(),
                        CH7  = calc.sheet.getCellByPosition(7, i).getFormula(),
                        CH8  = calc.sheet.getCellByPosition(8, i).getFormula(),
                        印加電圧    = Double.Parse(calc.sheet.getCellByPosition(9, i).getFormula()),
                        印加時間    = Double.Parse(calc.sheet.getCellByPosition(10, i).getFormula()),
                        絶縁抵抗値  = Double.Parse(calc.sheet.getCellByPosition(11, i).getFormula())
                    };
                    Parameter絶縁スペックBcr.Add(spec);
                    i++;
                }

                // sheetを取得
                calc.SelectSheet("SpecQ890");

                //耐電圧試験 規格値の読み出し
                splashMessage("Q890耐電圧試験 規格値のロード・・・");
                i = 2;//3行目から読み込む
                for (; ; )
                {
                    Application.DoEvents();
                    if (calc.sheet.getCellByPosition(0, i).getFormula() == "") break;
                    var spec = new 耐電圧試験スペック()
                    {
                        ステップ = calc.sheet.getCellByPosition(0, i).getFormula(),
                        CH1 = calc.sheet.getCellByPosition(1, i).getFormula(),
                        CH2 = calc.sheet.getCellByPosition(2, i).getFormula(),
                        CH3 = calc.sheet.getCellByPosition(3, i).getFormula(),
                        CH4 = calc.sheet.getCellByPosition(4, i).getFormula(),
                        CH5 = calc.sheet.getCellByPosition(5, i).getFormula(),
                        CH6 = calc.sheet.getCellByPosition(6, i).getFormula(),
                        CH7 = calc.sheet.getCellByPosition(7, i).getFormula(),
                        CH8 = calc.sheet.getCellByPosition(8, i).getFormula(),
                        印加電圧 = Double.Parse(calc.sheet.getCellByPosition(9, i).getFormula()),
                        印加時間 = Double.Parse(calc.sheet.getCellByPosition(10, i).getFormula()),
                        漏れ電流 = Double.Parse(calc.sheet.getCellByPosition(11, i).getFormula())
                    };
                    Parameter耐圧スペックQ890.Add(spec);
                    i++;
                }

                //絶縁抵抗試験 規格値の読み出し
                splashMessage("Q890絶縁抵抗試験 規格値のロード・・・");
                i = 10;//11行目から読み込む
                for (; ; )
                {
                    Application.DoEvents();
                    if (calc.sheet.getCellByPosition(0, i).getFormula() == "") break;
                    var spec = new 絶縁抵抗試験スペック()
                    {
                        ステップ = calc.sheet.getCellByPosition(0, i).getFormula(),
                        CH1 = calc.sheet.getCellByPosition(1, i).getFormula(),
                        CH2 = calc.sheet.getCellByPosition(2, i).getFormula(),
                        CH3 = calc.sheet.getCellByPosition(3, i).getFormula(),
                        CH4 = calc.sheet.getCellByPosition(4, i).getFormula(),
                        CH5 = calc.sheet.getCellByPosition(5, i).getFormula(),
                        CH6 = calc.sheet.getCellByPosition(6, i).getFormula(),
                        CH7 = calc.sheet.getCellByPosition(7, i).getFormula(),
                        CH8 = calc.sheet.getCellByPosition(8, i).getFormula(),
                        印加電圧 = Double.Parse(calc.sheet.getCellByPosition(9, i).getFormula()),
                        印加時間 = Double.Parse(calc.sheet.getCellByPosition(10, i).getFormula()),
                        絶縁抵抗値 = Double.Parse(calc.sheet.getCellByPosition(11, i).getFormula())
                    };
                    Parameter絶縁スペックQ890.Add(spec);
                    i++;
                }

                // sheetを取得
                calc.SelectSheet("SpecAur");

                //耐電圧試験 規格値の読み出し
                splashMessage("AUR耐電圧試験 規格値のロード・・・");
                i = 2;//3行目から読み込む
                for (; ; )
                {
                    Application.DoEvents();
                    if (calc.sheet.getCellByPosition(0, i).getFormula() == "") break;
                    var spec = new 耐電圧試験スペック()
                    {
                        ステップ = calc.sheet.getCellByPosition(0, i).getFormula(),
                        CH1 = calc.sheet.getCellByPosition(1, i).getFormula(),
                        CH2 = calc.sheet.getCellByPosition(2, i).getFormula(),
                        CH3 = calc.sheet.getCellByPosition(3, i).getFormula(),
                        CH4 = calc.sheet.getCellByPosition(4, i).getFormula(),
                        CH5 = calc.sheet.getCellByPosition(5, i).getFormula(),
                        CH6 = calc.sheet.getCellByPosition(6, i).getFormula(),
                        CH7 = calc.sheet.getCellByPosition(7, i).getFormula(),
                        CH8 = calc.sheet.getCellByPosition(8, i).getFormula(),
                        印加電圧 = Double.Parse(calc.sheet.getCellByPosition(9, i).getFormula()),
                        印加時間 = Double.Parse(calc.sheet.getCellByPosition(10, i).getFormula()),
                        漏れ電流 = Double.Parse(calc.sheet.getCellByPosition(11, i).getFormula())
                    };
                    Parameter耐圧スペックAUR.Add(spec);
                    i++;
                }

                //絶縁抵抗試験 規格値の読み出し
                splashMessage("AUR絶縁抵抗試験 規格値のロード・・・");
                i = 10;//11行目から読み込む
                for (; ; )
                {
                    Application.DoEvents();
                    if (calc.sheet.getCellByPosition(0, i).getFormula() == "") break;
                    var spec = new 絶縁抵抗試験スペック()
                    {
                        ステップ = calc.sheet.getCellByPosition(0, i).getFormula(),
                        CH1 = calc.sheet.getCellByPosition(1, i).getFormula(),
                        CH2 = calc.sheet.getCellByPosition(2, i).getFormula(),
                        CH3 = calc.sheet.getCellByPosition(3, i).getFormula(),
                        CH4 = calc.sheet.getCellByPosition(4, i).getFormula(),
                        CH5 = calc.sheet.getCellByPosition(5, i).getFormula(),
                        CH6 = calc.sheet.getCellByPosition(6, i).getFormula(),
                        CH7 = calc.sheet.getCellByPosition(7, i).getFormula(),
                        CH8 = calc.sheet.getCellByPosition(8, i).getFormula(),
                        印加電圧 = Double.Parse(calc.sheet.getCellByPosition(9, i).getFormula()),
                        印加時間 = Double.Parse(calc.sheet.getCellByPosition(10, i).getFormula()),
                        絶縁抵抗値 = Double.Parse(calc.sheet.getCellByPosition(11, i).getFormula())
                    };
                    Parameter絶縁スペックAUR.Add(spec);
                    i++;
                }


                // sheetを取得　"OperatorName"
                calc.SelectSheet("OperatorName");

                //作業者一覧の読み出し
                splashMessage("作業者一覧ののロード・・・");
                i = 1;//行インデックス
                for (; ; )
                {
                    Application.DoEvents();
                    string name = calc.sheet.getCellByPosition(1, i).getFormula(); if (name == "予約" || i == 11) break;
                    ParameterOperator.Add(name);
                    i++;
                }

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




   
    }
}
