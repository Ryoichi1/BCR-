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
    public partial class DailyCheckForm : Form
    {
        private bool TestFlag;


        耐電圧試験スペック BcrCh13 = new 耐電圧試験スペック()
        {
            CH1         = "2",  //High
            CH2         = "0",  //Open
            CH3         = "1",  //Low
            CH4         = "0",  //Open
            印加電圧    = 100,  //AC100V
            印加時間    = 1,    //1秒
            漏れ電流    = 1     //1mA
        };

        耐電圧試験スペック BcrCh24 = new 耐電圧試験スペック()
        {
            CH1         = "0",  //Open
            CH2         = "2",  //High
            CH3         = "0",  //Open
            CH4         = "1",  //Low
            印加電圧    = 100,  //AC100V
            印加時間    = 1,    //1秒
            漏れ電流    = 1     //1mAA
        };

        耐電圧試験スペック AurCh12 = new 耐電圧試験スペック()
        {
            CH1         = "2",  //High
            CH2         = "1",  //Low
            CH3         = "0",  //Open
            CH4         = "0",  //Open
            印加電圧    = 100,  //AC100V
            印加時間    = 1,    //1秒
            漏れ電流    = 1     //1mAA
        };   

        public DailyCheckForm()
        {
            InitializeComponent();
        }

        private void DailyCheckForm_Load(object sender, EventArgs e)
        {

            labelBcrCh1Ch3Check.ForeColor = Color.Black;
            labelBcrCh2Ch4Check.ForeColor = Color.Black;
            labelAurCh1Ch2Check.ForeColor = Color.Black;

            labelDecision.Text = "";
            labelErrorMess.Text = "";

            labelDanger.BackColor = Color.MistyRose;

            //タイマーの設定
            timerLbMessage.Interval = 500;
            timerLbMessage.Enabled = true;

            timerCount.Interval = 1000;
            timerCount.Stop();
            
            pictureBox1.ImageLocation = Constants.SetBcrPath;
            labelMessage.Text = "ＢＣＲ側の治具に点検用サンプルをセットして開始ボタンを押してください";

            
        }

        private void buttonStart_Click(object sender, EventArgs e)
        {
            TestFlag = true;
            
            timerLbMessage.Enabled = false;
            buttonReturn.Enabled = false;
            buttonStart.Enabled = false;
            buttonStart.BackColor = Color.Transparent;

            if (State.OperatorName == "")
            {
                MessageBox.Show("作業者名が選択されていません！");
                this.Close();
                return;
            }

            //BCR側の治具にサンプルがセットされているかチェック
            Method.io.ReadInputData(EPX64R.PortName.P7);
            if ((byte)(Method.io.P7InputData & 0x03) != 0x02)
            {
                MessageBox.Show("ＢＣＲサンプルが正しくセットされていません");
                timerLbMessage.Enabled = true;

                while (true)
                {
                    Application.DoEvents();
                    Method.io.ReadInputData(EPX64R.PortName.P7);
                    if ((byte)(Method.io.P7InputData & 0x03) != 0x03)
                    {
                        labelMessage.Text = "点検用サンプルを取り外してください";
                    }
                    else break;
                }
                
               
                buttonReturn.Enabled = true;
                buttonStart.Enabled = true;
                labelMessage.Text = "ＢＣＲ側の治具に点検用サンプルをセットして開始ボタンを押してください";
                pictureBox1.ImageLocation = Constants.SetBcrPath;
                return;

            }

            //オペレーターへ電圧印加の警告をする
            labelMessage.BackColor = Color.Yellow;

            State.Count = 5;
            Flags.FlagCount = false;
            timerCount.Start();

            for (; ; )
            {
                Application.DoEvents();
                if (Flags.FlagCount) break;
                labelMessage.Text = Constants.MessWarning + State.Count.ToString() + "秒";
                labelMessage.Refresh();
            }



            labelMessage.BackColor = Color.LightBlue;
            labelMessage.Text = "点検中・・・";

            labelDanger.BackColor = Color.Red;
            labelDecision.Text = "";
            labelBcrCh1Ch3Check.ForeColor = Color.Black;
            labelBcrCh2Ch4Check.ForeColor = Color.Black;

            Method.InterLock解除();


            TOS9200.SendCommand("STOP", false);
            TOS9200.SendQuery("PROT?");
            TOS9200.ReadRecieveData(500);
            if (TOS9200.RecieveData != "0,0")
            {
                MessageBox.Show("危険区域に入らないでください！");
                goto StepNg;
            }
            //ｱｸﾉﾘｯｼﾞﾒｯｾｰｼﾞを返す設定
            if(!TOS9200.SendCommand("SIL 0")) goto StepNg;


            //ＣＨ１－３の導通チェック●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●
            TOS9200.SendCommand("STOP");
            TOS9200.Dsr = 0;    //DSRレジスタ初期化
            TOS9200.Fail = 0;   //failレジスタ初期化
            TOS9200.SetFunction(BcrCh13, 1,"ON");//試験設定

            TOS9200.TosStart();

            if(TOS9200.ErrorMess != TOS9200.ErrorMessage.漏洩電流上限異常)
            {
                labelBcrCh1Ch3Check.ForeColor = Color.Red;
                goto StepNg;
            }

            TOS9200.SendCommand("STOP");
            labelBcrCh1Ch3Check.ForeColor = Color.LightSeaGreen;


            //ＣＨ２－４の導通チェック●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●
            TOS9200.Dsr = 0;    //DSRレジスタ初期化
            TOS9200.Fail = 0;   //failレジスタ初期化
            TOS9200.SetFunction(BcrCh24, 1, "ON");//試験設定

            TOS9200.TosStart();

            if (TOS9200.ErrorMess != TOS9200.ErrorMessage.漏洩電流上限異常)
            {
                labelBcrCh2Ch4Check.ForeColor = Color.Red;
                goto StepNg;
            }

            TOS9200.SendCommand("STOP");
            labelBcrCh2Ch4Check.ForeColor = Color.LightSeaGreen;


            labelDanger.BackColor = Color.MistyRose;

            Method.InterLock発動();

            //ＡＵＲ　ＣＨ１－２の導通チェック●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●
            pictureBox1.ImageLocation = Constants.SetAurPath;

            labelMessage.Text = "ＡＵＲ側の治具に点検用サンプルをセットしてください";
            MessageBox.Show("ＡＵＲ側の治具に点検用サンプルをセットしてください");

            //BCR側の治具にサンプルがセットされているかチェック
            Method.io.ReadInputData(EPX64R.PortName.P7);
            if ((byte)(Method.io.P7InputData & 0x03) != 0x01)
            {
                MessageBox.Show("ＡＵＲサンプルが正しくセットされていません");
                goto StepNg;
            }

            //オペレーターへ電圧印加の警告をする
            labelMessage.BackColor = Color.Yellow;

            State.Count = 5;
            Flags.FlagCount = false;
            timerCount.Start();

            for (; ; )
            {
                Application.DoEvents();
                if (Flags.FlagCount) break;
                labelMessage.Text = Constants.MessWarning + State.Count.ToString() + "秒";
                labelMessage.Refresh();
            }



            labelMessage.BackColor = Color.LightBlue;
            labelMessage.Text = "点検中・・・";

            Method.InterLock解除();

            TOS9200.SendCommand("STOP");
            TOS9200.Dsr = 0;    //DSRレジスタ初期化
            TOS9200.Fail = 0;   //failレジスタ初期化
            TOS9200.SetFunction(AurCh12, 1, "ON");//試験設定

            TOS9200.TosStart();

            if (TOS9200.ErrorMess != TOS9200.ErrorMessage.漏洩電流上限異常)
            {
                labelAurCh1Ch2Check.ForeColor = Color.Red;
                goto StepNg;
            }

            TOS9200.SendCommand("STOP");
            labelAurCh1Ch2Check.ForeColor = Color.LightSeaGreen;


            //日常点検合格の処理●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●
            Method.InterLock発動();

            if (!SaveDailyCheck())
            {
                MessageBox.Show("試験データの保存に失敗しました");
                this.Close();
                return;
            }
            labelDanger.BackColor = Color.SeaShell;
            timerLbMessage.Enabled = true;
            labelDecision.Text = "PASS";
            labelDecision.ForeColor = Color.LightSeaGreen;
            
            while (true)
            {
                Application.DoEvents();
                Method.io.ReadInputData(EPX64R.PortName.P7);
                if ((byte)(Method.io.P7InputData & 0x03) != 0x03)
                {
                    labelMessage.Text = "点検用サンプルを取り外してください";
                }
                else break;
            }
            
            
            
            this.Close();

            return;

        StepNg://日常点検不合格の処理●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●

            TOS9200.SendCommand("STOP");
            labelDanger.BackColor = Color.SeaShell;

            timerLbMessage.Enabled = true;
            labelDecision.Text = "FAIL";
            labelDecision.ForeColor = Color.Red;

            //エラーメッセージの生成
            Func<string> GetErrorMess = () =>
            {
                switch (TOS9200.ErrorMess)
                {
                    case TOS9200.ErrorMessage.正常終了:
                        return "ケーブル断線";
                    case TOS9200.ErrorMessage.強制停止:
                        return "強制停止";
                    case TOS9200.ErrorMessage.インターロック作動:
                        return "インターロック作動";
                    case TOS9200.ErrorMessage.コンタクトエラー:
                        return "コンタクトエラー";
                    default:
                        return "";
                }
            };
            labelErrorMess.Text = GetErrorMess();
            
            Method.InterLock発動();


            while (true)
            {
                Application.DoEvents();
                Method.io.ReadInputData(EPX64R.PortName.P7);
                if ((byte)(Method.io.P7InputData & 0x03) != 0x03)
                {
                    labelMessage.Text = "点検用サンプルを取り外してください";
                }
                else break;
            }
 




            labelMessage.Text = "ＢＣＲ側の治具に点検用サンプルをセットして開始ボタンを押してください";
            pictureBox1.ImageLocation = Constants.SetBcrPath;
            labelBcrCh1Ch3Check.ForeColor = Color.Black;
            labelBcrCh2Ch4Check.ForeColor = Color.Black;
            labelAurCh1Ch2Check.ForeColor = Color.Black;
            labelDecision.Text = "";
            labelErrorMess.Text = "";
            buttonReturn.Enabled = true;
            buttonStart.Enabled = true;
            TestFlag = false;
            return;
        }



        private void buttonReturn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public bool SaveDailyCheck()
        {
            OpenOffice calc = new OpenOffice();
            
            try
            {
                //当月点検データファイルが存在するかどうかの判定
                string date = DateTime.Now.ToString("yyyy年MM月dd日 HH:mm:ss"); //dateは"yyyy年MM月dd日 HH:mm:ss"の形式で保存
                string year = date.Substring(0, 4);
                string month = date.Substring(5, 2);

                if (!System.IO.File.Exists(Constants.日常点検FolderPath + year + "年" + month + "月.ods"))
                {
                    //既存検査データがなければ新規作成
                    File.Copy(Constants.日常点検FolderPath + "format.ods", Constants.日常点検FolderPath + year + "年" + month + "月.ods");
                }

                //当月日常点検ファイルを開く
                string filepath = Constants.日常点検FolderPath + year + "年" + month + "月.ods";
                calc.OpenFile(filepath);

                // sheetを取得
                calc.SelectSheet("Sheet1");

                int newRow = 0;
                for (; ; )//使用されているセルの最終行を検索する
                {
                    Application.DoEvents();
                    if (calc.sheet.getCellByPosition(0, newRow).getFormula() == "") break;
                    newRow++;
                }

                //最終行にデータを追加
                int i = 0;
                foreach (var data in new string[] { date, "OK", State.OperatorName })
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
                calc.CloseFile();//calcファイルが開いている場合のみクローズ処理する
            }
        }

        private void timerLbMessage_Tick(object sender, EventArgs e)
        {
            if (labelMessage.BackColor == Color.Transparent)
            {
                labelMessage.BackColor = Color.LightBlue;
                if (!TestFlag) buttonStart.BackColor = Color.LightBlue; 
            }
            else
            {
                labelMessage.BackColor = Color.Transparent;
                buttonStart.BackColor = Color.Transparent;
            }     
        }

        private void timerCount_Tick(object sender, EventArgs e)
        {
            if (State.Count > 0)
            {
                State.Count--;
            }
            else
            {
                timerCount.Stop();
                Flags.FlagCount = true;
            }
        }


    
    
    }
}
