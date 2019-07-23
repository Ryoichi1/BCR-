using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace BCR耐電圧_絶縁抵抗試験
{
    public partial class SetSpecForm : Form
    {

        double 耐圧1印加電圧;
        double 耐圧1印加時間;
        double 耐圧1漏れ電流;

        double 絶縁1印加電圧;
        double 絶縁1印加時間;
        double 絶縁1抵抗値;

        double 絶縁2印加電圧;
        double 絶縁2印加時間;
        double 絶縁2抵抗値;

        double 絶縁3印加電圧;
        double 絶縁3印加時間;
        double 絶縁3抵抗値;


        public SetSpecForm()
        {
            InitializeComponent();
        }

        private void SetSpecForm_Load(object sender, EventArgs e)
        {
            //耐電圧試験項目のセット
            comboBox耐圧1_一次端子.Text   = SetContact(State.耐電圧スペックリストBcr[0].CH1);
            comboBox耐圧1_二次端子.Text   = SetContact(State.耐電圧スペックリストBcr[0].CH2);
            comboBox耐圧1_ダンパ端子.Text = SetContact(State.耐電圧スペックリストBcr[0].CH3);
            comboBox耐圧1_モニタ端子.Text = SetContact(State.耐電圧スペックリストBcr[0].CH4);
            textBox耐圧1印加電圧.Text     = State.耐電圧スペックリストBcr[0].印加電圧.ToString("F0");
            textBox耐圧1印加時間.Text     = State.耐電圧スペックリストBcr[0].印加時間.ToString("F0");
            textBox耐圧1漏れ電流.Text     = State.耐電圧スペックリストBcr[0].漏れ電流.ToString("F0");


            //絶縁抵抗試験項目のセット
            comboBox絶縁1_一次端子.Text   = SetContact(State.絶縁抵抗スペックリストBcr[0].CH1);
            comboBox絶縁1_二次端子.Text   = SetContact(State.絶縁抵抗スペックリストBcr[0].CH2);
            comboBox絶縁1_ダンパ端子.Text = SetContact(State.絶縁抵抗スペックリストBcr[0].CH3);
            comboBox絶縁1_モニタ端子.Text = SetContact(State.絶縁抵抗スペックリストBcr[0].CH4);
            textBox絶縁1印加電圧.Text     = State.絶縁抵抗スペックリストBcr[0].印加電圧.ToString("F0");
            textBox絶縁1印加時間.Text     = State.絶縁抵抗スペックリストBcr[0].印加時間.ToString("F0");
            textBox絶縁1抵抗値.Text       = State.絶縁抵抗スペックリストBcr[0].絶縁抵抗値.ToString("F0");

            comboBox絶縁2_一次端子.Text   = SetContact(State.絶縁抵抗スペックリストBcr[1].CH1);
            comboBox絶縁2_二次端子.Text   = SetContact(State.絶縁抵抗スペックリストBcr[1].CH2);
            comboBox絶縁2_ダンパ端子.Text = SetContact(State.絶縁抵抗スペックリストBcr[1].CH3);
            comboBox絶縁2_モニタ端子.Text = SetContact(State.絶縁抵抗スペックリストBcr[1].CH4);
            textBox絶縁2印加電圧.Text     = State.絶縁抵抗スペックリストBcr[1].印加電圧.ToString("F0");
            textBox絶縁2印加時間.Text     = State.絶縁抵抗スペックリストBcr[1].印加時間.ToString("F0");
            textBox絶縁2抵抗値.Text       = State.絶縁抵抗スペックリストBcr[1].絶縁抵抗値.ToString("F0");

            comboBox絶縁3_一次端子.Text   = SetContact(State.絶縁抵抗スペックリストBcr[2].CH1);
            comboBox絶縁3_二次端子.Text   = SetContact(State.絶縁抵抗スペックリストBcr[2].CH2);
            comboBox絶縁3_ダンパ端子.Text = SetContact(State.絶縁抵抗スペックリストBcr[2].CH3);
            comboBox絶縁3_モニタ端子.Text = SetContact(State.絶縁抵抗スペックリストBcr[2].CH4);
            textBox絶縁3印加電圧.Text     = State.絶縁抵抗スペックリストBcr[2].印加電圧.ToString("F0");
            textBox絶縁3印加時間.Text     = State.絶縁抵抗スペックリストBcr[2].印加時間.ToString("F0");
            textBox絶縁3抵抗値.Text       = State.絶縁抵抗スペックリストBcr[2].絶縁抵抗値.ToString("F0");


        }



        public bool CheckSpec()
        {


            //設定値が正しく入力されているかの確認

            //耐電圧STEP1 印加電圧値のチェック
            if (!Double.TryParse(textBox耐圧1印加電圧.Text, out 耐圧1印加電圧))
            {
                MessageBox.Show("耐電圧 STEP1の印加電圧が正しく入力されていません");
                textBox耐圧1印加電圧.Text = State.耐電圧スペックリストBcr[0].印加電圧.ToString("F0");　//デフォルト値に戻す
                textBox耐圧1印加電圧.Focus();
                return false;
            }


            //耐電圧STEP1 印加時間のチェック
            if (!Double.TryParse(textBox耐圧1印加時間.Text, out 耐圧1印加時間))
            {
                MessageBox.Show("耐電圧 STEP1の印加時間が正しく入力されていません");
                textBox耐圧1印加時間.Text = State.耐電圧スペックリストBcr[0].印加時間.ToString("F0");　//デフォルト値に戻す
                textBox耐圧1印加時間.Focus();
                return false;
            }


            //耐電圧STEP1 漏れ電流のチェック
            if (!Double.TryParse(textBox耐圧1漏れ電流.Text, out 耐圧1漏れ電流))
            {
                MessageBox.Show("耐電圧 STEP1の漏れ電流が正しく入力されていません");
                textBox耐圧1漏れ電流.Text = State.耐電圧スペックリストBcr[0].漏れ電流.ToString("F0");　//デフォルト値に戻す
                textBox耐圧1漏れ電流.Focus();
                return false;
            }



//*********************************************************************************************************************

            //絶縁抵抗STEP1 印加電圧値のチェック
            if (!Double.TryParse(textBox絶縁1印加電圧.Text, out 絶縁1印加電圧))
            {
                MessageBox.Show("絶縁 STEP1の印加電圧が正しく入力されていません");
                textBox絶縁1印加電圧.Text = State.絶縁抵抗スペックリストBcr[0].印加電圧.ToString("F0");　//デフォルト値に戻す
                textBox絶縁1印加電圧.Focus();
                return false;
            }

            //絶縁抵抗STEP2 印加電圧値のチェック
            if (!Double.TryParse(textBox絶縁2印加電圧.Text, out 絶縁2印加電圧))
            {
                MessageBox.Show("絶縁 STEP2の印加電圧が正しく入力されていません");
                textBox絶縁2印加電圧.Text = State.絶縁抵抗スペックリストBcr[1].印加電圧.ToString("F0");　//デフォルト値に戻す
                textBox絶縁2印加電圧.Focus();
                return false;
            }

            //絶縁抵抗STEP3 印加電圧値のチェック
            if (!Double.TryParse(textBox絶縁3印加電圧.Text, out 絶縁3印加電圧))
            {
                MessageBox.Show("絶縁 STEP3の印加電圧が正しく入力されていません");
                textBox絶縁3印加電圧.Text = State.絶縁抵抗スペックリストBcr[2].印加電圧.ToString("F0");　//デフォルト値に戻す
                textBox絶縁3印加電圧.Focus();
                return false;
            }

            //絶縁抵抗STEP1 印加時間のチェック
            if (!Double.TryParse(textBox絶縁1印加時間.Text, out 絶縁1印加時間))
            {
                MessageBox.Show("絶縁 STEP1の印加時間が正しく入力されていません");
                textBox絶縁1印加時間.Text = State.絶縁抵抗スペックリストBcr[0].印加時間.ToString("F0");　//デフォルト値に戻す
                textBox絶縁1印加時間.Focus();
                return false;
            }

            //絶縁抵抗STEP2 印加時間のチェック
            if (!Double.TryParse(textBox絶縁2印加時間.Text, out 絶縁2印加時間))
            {
                MessageBox.Show("絶縁 STEP2の印加時間が正しく入力されていません");
                textBox絶縁2印加時間.Text = State.絶縁抵抗スペックリストBcr[1].印加時間.ToString("F0");　//デフォルト値に戻す
                textBox絶縁2印加時間.Focus();
                return false;
            }

            //絶縁抵抗STEP3 印加時間のチェック
            if (!Double.TryParse(textBox絶縁3印加時間.Text, out 絶縁3印加時間))
            {
                MessageBox.Show("絶縁 STEP3の印加時間が正しく入力されていません");
                textBox絶縁3印加時間.Text = State.絶縁抵抗スペックリストBcr[2].印加時間.ToString("F0");　//デフォルト値に戻す
                textBox絶縁3印加時間.Focus();
                return false;
            }

            //絶縁抵抗STEP1 抵抗値のチェック
            if (!Double.TryParse(textBox絶縁1抵抗値.Text, out 絶縁1抵抗値))
            {
                MessageBox.Show("絶縁 STEP1の絶縁抵抗値が正しく入力されていません");
                textBox絶縁1抵抗値.Text = State.絶縁抵抗スペックリストBcr[0].絶縁抵抗値.ToString("F0");　//デフォルト値に戻す
                textBox絶縁1抵抗値.Focus();
                return false;
            }

            //絶縁抵抗STEP2 抵抗値のチェック
            if (!Double.TryParse(textBox絶縁2抵抗値.Text, out 絶縁2抵抗値))
            {
                MessageBox.Show("絶縁 STEP2の絶縁抵抗値が正しく入力されていません");
                textBox絶縁2抵抗値.Text = State.絶縁抵抗スペックリストBcr[1].絶縁抵抗値.ToString("F0");　//デフォルト値に戻す
                textBox絶縁2抵抗値.Focus();
                return false;
            }

            //絶縁抵抗STEP3 抵抗値のチェック
            if (!Double.TryParse(textBox絶縁3抵抗値.Text, out 絶縁3抵抗値))
            {
                MessageBox.Show("絶縁 STEP3の絶縁抵抗値が正しく入力されていません");
                textBox絶縁3抵抗値.Text = State.絶縁抵抗スペックリストBcr[2].絶縁抵抗値.ToString("F0");　//デフォルト値に戻す
                textBox絶縁3抵抗値.Focus();
                return false;
            }
            
            return true;
        
        
        }

        private void buttonOk_Click(object sender, EventArgs e)
        {
            buttonOk.Enabled = false;
            if (!CheckSpec())
            {
                buttonOk.Enabled = true;
                return;
            }

            //パラメータファイルの更新
            var 耐圧スペックリスト = new List<耐電圧試験スペック>();

            耐圧スペックリスト.Add(new 耐電圧試験スペック() { 
                CH1    = ResetContact(comboBox耐圧1_一次端子.Text),
                CH2    = ResetContact(comboBox耐圧1_二次端子.Text),
                CH3  = ResetContact(comboBox耐圧1_ダンパ端子.Text),
                CH4  = ResetContact(comboBox耐圧1_モニタ端子.Text),
                印加電圧 　 = 耐圧1印加電圧,
                印加時間 　 = 耐圧1印加時間,
                漏れ電流 　 = 耐圧1漏れ電流
            });


            var 絶縁スペックリスト = new List<絶縁抵抗試験スペック>();

            絶縁スペックリスト.Add(new 絶縁抵抗試験スペック()
            {
                CH1    = ResetContact(comboBox絶縁1_一次端子.Text),
                CH2    = ResetContact(comboBox絶縁1_二次端子.Text),
                CH3  = ResetContact(comboBox絶縁1_ダンパ端子.Text),
                CH4  = ResetContact(comboBox絶縁1_モニタ端子.Text),
                印加電圧 　 = 絶縁1印加電圧,
                印加時間 　 = 絶縁1印加時間,
                絶縁抵抗値  = 絶縁1抵抗値
            });

            絶縁スペックリスト.Add(new 絶縁抵抗試験スペック()
            {
                CH1 = ResetContact(comboBox絶縁2_一次端子.Text),
                CH2 = ResetContact(comboBox絶縁2_二次端子.Text),
                CH3 = ResetContact(comboBox絶縁2_ダンパ端子.Text),
                CH4 = ResetContact(comboBox絶縁2_モニタ端子.Text),
                印加電圧 = 絶縁2印加電圧,
                印加時間 = 絶縁2印加時間,
                絶縁抵抗値 = 絶縁2抵抗値
            });

            絶縁スペックリスト.Add(new 絶縁抵抗試験スペック()
            {
                CH1 = ResetContact(comboBox絶縁3_一次端子.Text),
                CH2 = ResetContact(comboBox絶縁3_二次端子.Text),
                CH3 = ResetContact(comboBox絶縁3_ダンパ端子.Text),
                CH4 = ResetContact(comboBox絶縁3_モニタ端子.Text),
                印加電圧 = 絶縁3印加電圧,
                印加時間 = 絶縁3印加時間,
                絶縁抵抗値 = 絶縁3抵抗値
            });



            if (!SaveSpec(耐圧スペックリスト, 絶縁スペックリスト))
            {
                MessageBox.Show("パラメータファイル読み出し異常\r\n" + "アプリケーションを終了します");
                Environment.Exit(0);
            }

            //自動でスペック変更フォームを閉じる
            this.Close();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            buttonCancel.Enabled = false;
            this.Close();
        }

        //**************************************************************************
        //
        //引数：
        //戻値：
        //**************************************************************************
        private string SetContact(string data)
        {
            switch (data)
            {
                case "0":
                    return "Open";
                case "1":
                    return "Low";
                case "2":
                    return "High";
                default:
                    return "Open";
            }
        }
        //**************************************************************************
        //
        //引数：
        //戻値：
        //**************************************************************************
        private string ResetContact(string data)
        {
            switch (data)
            {
                case "Open":
                    return "0";
                case "Low":
                    return "1";
                case "High":
                    return "2";
                default:
                    return "0";
            }
        }

        //**************************************************************************
        //試験パラメータの編集
        //引数：
        //戻値：
        //**************************************************************************
        public bool SaveSpec(List<耐電圧試験スペック> 耐圧スペック, List<絶縁抵抗試験スペック> 絶縁スペック)
        {
            OpenOffice calc = new OpenOffice();

            try
            {
                //parameterファイルを開く
                calc.OpenFile(Constants.ParameterFilePath);

                // sheetを取得
                calc.SelectSheet("Spec");

                //①耐電圧試験のスペック更新               
                int i = 2;//parameterファイルの3行目から更新
                foreach(var data in 耐圧スペック)　
                {
                    calc.sheet.getCellByPosition(1, i).setFormula(data.CH1);
                    calc.sheet.getCellByPosition(2, i).setFormula(data.CH2);
                    calc.sheet.getCellByPosition(3, i).setFormula(data.CH3);
                    calc.sheet.getCellByPosition(4, i).setFormula(data.CH4);
                    calc.sheet.getCellByPosition(5, i).setFormula(data.印加電圧.ToString("F0"));
                    calc.sheet.getCellByPosition(6, i).setFormula(data.印加時間.ToString("F1"));
                    calc.sheet.getCellByPosition(7, i).setFormula(data.漏れ電流.ToString("F0"));
                    i++;
                }

                //②絶縁抵抗試験験のスペック更新               
                i = 10;//parameterファイルの11行目から更新
                foreach (var data in 絶縁スペック)　
                {
                    calc.sheet.getCellByPosition(1, i).setFormula(data.CH1);
                    calc.sheet.getCellByPosition(2, i).setFormula(data.CH2);
                    calc.sheet.getCellByPosition(3, i).setFormula(data.CH3);
                    calc.sheet.getCellByPosition(4, i).setFormula(data.CH4);
                    calc.sheet.getCellByPosition(5, i).setFormula(data.印加電圧.ToString("F0"));
                    calc.sheet.getCellByPosition(6, i).setFormula(data.印加時間.ToString("F1"));
                    calc.sheet.getCellByPosition(7, i).setFormula(data.絶縁抵抗値.ToString("F0"));
                    i++;
                }

                // Calcファイルを保存して閉じる
                if (!calc.SaveFile())
                {
                    return false;
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
