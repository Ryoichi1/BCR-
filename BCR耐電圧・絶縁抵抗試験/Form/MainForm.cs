using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BCR耐電圧_絶縁抵抗試験
{
    public partial class MainForm : Form
    {



        public MainForm()
        {
            InitializeComponent();
        }


        //フォームイベント
        private void MainForm_Load(object sender, EventArgs e)
        {
            //IOボードの初期化
            Method.io.InitEpx64r(0x10);// 0x10→ P7:0 P6:0 P5:0 P4:1 P3:0 P2:0 P1:0 P0:0

            Method.ResetIo();

            //タイマーの初期化
            timerCurrentTime.Interval = 30;
            timerCurrentTime.Enabled = true;

            timerLbMessage.Interval = 1000;
            timerLbMessage.Enabled = true;

            timerTextInput.Interval = 800;
            timerTextInput.Stop();

            timerCount.Interval = 1000;
            timerCount.Stop();

            timerLed.Interval = 1000;
            timerLed.Stop();

            //Flagsクラスへのデリゲート設定
            Flags.日常点検ラベルのセット = (flag) =>
            {
                labelDailyCheck.BackColor = (flag) ? Color.MediumSeaGreen : SystemColors.ActiveBorder;
            };

            //ＰＣ１へアクセス可能か確認する
            if (!System.IO.Directory.Exists(Constants.Pc3LogFileFolder))
            {
                MessageBox.Show("ＰＣ１へアクセスできません\r\n" + "ＰＣ１を再起動してください\r\n" + "アプリケーションを終了します");
                Environment.Exit(0);
            }


            //TOS9200の初期化
            SplashForm._form.SetLabel("TOS9200初期化中・・・");
            TOS9200.InitTOS9200(PORT_NAME.COM1);
            if (!TOS9200.Query("*IDN?", (data) => data.IndexOf("KIKUSUI") >= 0))
            {
                MessageBox.Show("TOS9200通信異常\r\n" + "アプリケーションを終了します");
                Environment.Exit(0);
            }

            //パラメータファイルの読み出し
            SplashForm._form.SetLabel("パラメータファイル読み出し中・・・");
            if (!State.LoadParameter())
            {
                MessageBox.Show("パラメータファイルロード異常\r\n" + "アプリケーションを終了します");
                Environment.Exit(0);
            }

            //フォームロード時、comboBoxOperatorNameをアクティブにする
            this.ActiveControl = comboBoxOperatorName;

            //日常点検を実施しているかどうか、フラグを設定する
            SplashForm._form.SetLabel("日常点検実施チェック・・・");
            Method.SetDailyCheckFlg();

            //ライトカーテンに電源供給する
            Method.io.OutBit(EPX64R.PortName.P4, EPX64R.BitName.b0, EPX64R.OutData.H);
            General.Wait(1500);
            //ＩＯユニットの電源スイッチがＯＮになっているかのチェック
            Method.io.ReadInputData(EPX64R.PortName.P7);
            if ((byte)(Method.io.P7InputData & 0x08) != 0x08)
            {
                MessageBox.Show("ＩＯユニットの電源スイッチがＯＮになっていません！");
                Method.ResetIo();
                Environment.Exit(0);
            }
            //フォームの初期化
            InitForm();
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            //COMポートを閉じる
            TOS9200.CloseComPort();

            //ioリセット
            Method.ResetIo();

            //Openoffice（calc）のプロセス強制終了
            General.KillOpenOffice();

            MessageBox.Show("ＩＯユニットの電源をＯＦＦしてください！");
        }

        private void MainForm_Click(object sender, EventArgs e)
        {
            SetFocus();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)//耐圧試験機が動作中にフォームを閉じないようにする処理
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                if (Flags.FlagTest) e.Cancel = true;
            }
        }

        //ツールストリップメニューイベント
        private void MenuItemFile_Close_Click(object sender, EventArgs e)
        {
            Environment.Exit(0);
        }

        private void MenuItemUtility_SelfCheck_Click(object sender, EventArgs e)
        {
            timerLbMessage.Stop();
            panelOperator.BackColor = Color.Transparent;
            panelSerial.BackColor = Color.Transparent;
            labelMessage.BackColor = Color.Transparent;
            labelMessage.Text = "";

            this.Enabled = false;

            //日常点検フォームの作成
            var dForm = new DailyCheckForm();
            // Form をモーダルで表示する
            dForm.ShowDialog();

            // 不要になった時点で破棄する (正しくは オブジェクトの破棄を保証する を参照)
            dForm.Dispose();
            this.Refresh();

            Method.SetDailyCheckFlg();
            this.Enabled = true;

            timerLbMessage.Start();

            SetFocus();
        }



        private void MenuItemUtility_SetOperator_Click(object sender, EventArgs e)
        {
            timerLbMessage.Stop();
            panelOperator.BackColor = Color.Transparent;
            panelSerial.BackColor = Color.Transparent;
            labelMessage.BackColor = Color.Transparent;
            labelMessage.Text = "";

            this.Enabled = false;

            //アイテム追加フォームの作成
            var oForm = new SetOperatorForm();

            // Form をモーダルで表示する
            oForm.ShowDialog();

            // 不要になった時点で破棄する (正しくは オブジェクトの破棄を保証する を参照)
            oForm.Dispose();
            this.Refresh();


            //検査用パラメータのロード
            if (!State.LoadParameter())
            {
                MessageBox.Show("パラメータロード異常\r\n" + "アプリケーションを終了します");
                Environment.Exit(0);
            }

            this.Enabled = true;
            timerLbMessage.Start();
            //メインフォームの初期化
            InitForm();

            this.Refresh();
        }

        private void MenuItemHelp_Manual_Click(object sender, EventArgs e)
        {
            //ファイルを開いて終了まで待機する
            System.Diagnostics.Process p =
                System.Diagnostics.Process.Start(@"C:\BCR耐電圧・絶縁抵抗試験\BCR耐圧試験手順書.pdf");
            p.WaitForExit();
        }

        private void MenuItemHelp_VerInfo_Click(object sender, EventArgs e)
        {
            // Form1 の新しいインスタンスを生成する
            VerInfoForm vForm = new VerInfoForm();

            // Form1 をモーダルで表示する
            vForm.ShowDialog();

            // 不要になった時点で破棄する (正しくは オブジェクトの破棄を保証する を参照)
            vForm.Dispose();
        }


        //ボタンイベント
        private void buttonClearOperator_Click(object sender, EventArgs e)
        {
            comboBoxOperatorName.SelectedIndex = -1;//コンボボックスを空白にする
            作業者名 = false;
        }


        //テキストボックスイベント
        private void textBoxSerial_KeyPress(object sender, KeyPressEventArgs e)
        {

        }

        private void textBoxSerial_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            try
            {

                if (e.KeyCode == Keys.Tab)//QRコードリーダーにはサフィックス（TAB）の設定が必要です
                {
                    e.IsInputKey = true;
                    timerTextInput.Stop();

                    //********************************************************************************************************************
                    //QRコード → 型式,INDEXｺｰﾄﾞ
                    //BC-R20B1G0500,130425002　※INDEXコードの先頭"FCM"が西暦下2桁に変わっている、シリアル部は3桁
                    //AUR455C422500,190715029　※INDEXコードの先頭"FCM"が西暦下2桁に変わっている、シリアル部は3桁

                    //ﾛｸﾞﾃﾞｰﾀ →　INDEXｺｰﾄﾞ,型式,・・・　
                    //FCM04250002,BC-R20B1G0500, ※INDEXコードのシリアル部は4桁
                    //********************************************************************************************************************

                    //型式の抽出
                    var buf = textBoxSerial.Text.Split(',');
                    State.model = buf[0];
                    State.年月 = buf[1].Substring(0, 4);

                    //カテゴリ選択 & タブコントロールの切り替え
                    if (State.model.Contains("Q890"))
                    {
                        State.Category = State.CATEGORY.Q890;
                        tabControlBcr.SelectedTab = tabPage2;
                    }
                    else if (State.model.Contains("AUR455") || State.model.Contains("AUR355"))
                    {
                        State.Category = State.CATEGORY.NEW_AUR;
                        tabControlBcr.SelectedTab = tabPage3;
                    }
                    else
                    {
                        State.Category = State.CATEGORY.BCR;
                        tabControlBcr.SelectedTab = tabPage1;
                    }

                    //ログデータ内で検索する文字列を生成
                    State.index = "FCM"+ (buf[1].Insert(6, "0")).Substring(2);//西暦2桁をFCMに変換、連番部を3桁→4桁に変換

                    string index_model = State.index + "," + State.model;//ログデータ内で検索する文字列

                    //BCR or AUR455/355で場合分け
                    if (State.Category == State.CATEGORY.NEW_AUR)
                    {
                        State.pc3LogFilePath/*配線ラベル印刷機*/ = Constants.Aur_Pc3LogFileFolder + "BCR_PC3_" + State.年月 + ".txt";
                        State.pc4LogFilePath/*耐圧試験機*/ = Constants.Aur_Pc4LogFileFolder + "BCR_PC4_" + State.年月 + ".txt";
                    }
                    else
                    {
                        State.pc3LogFilePath/*配線ラベル印刷機*/ = Constants.Pc3LogFileFolder + "BCR_PC3_" + State.年月 + ".txt";
                        State.pc4LogFilePath/*耐圧試験機*/ = Constants.Pc4LogFileFolder + "BCR_PC4_" + State.年月 + ".txt";
                        State.pc7LogFilePath/*誤配線チェッカー*/ = Constants.Pc7LogFileFolder + "bcr_pc7_" + State.年月 + ".txt";
                    }


                    //ＰＣ１へアクセス可能か確認する（試験毎に確認する）
                    if (!System.IO.Directory.Exists(Constants.Pc3LogFileFolder))
                    {
                        MessageBox.Show("ＰＣ１へアクセスできません\r\n" + "ＰＣ１を再起動してください\r\n" + "アプリケーションを終了します");
                        timerTextInput.Start();
                        シリアル = false;
                        return;
                    }

                    //ログファイルチェック
                    //引数で渡したログファイル内に指定文字列（INDEXｺｰﾄﾞ,型式）があるかどうかの判定

                    //耐圧試験機のPCがVs2010なので、ローカル関数は使えない　このままにしておく
                    Func<string, bool> CheckLogFile = (LogFilePath) =>
                    {
                        using (var fs = new FileStream(
                            LogFilePath,
                            FileMode.Open,
                            FileAccess.Read,
                            FileShare.ReadWrite
                            ))
                        using (StreamReader sr = new StreamReader(fs, Encoding.GetEncoding("Shift_JIS")))
                        {
                            var list = new List<string>();
                            // ファイルの内容を1行ずつ読み込みます。
                            // 読み込んだ文字列の改行文字は削除されるので注意してください。
                            string line = null;
                            while ((line = sr.ReadLine()) != null)
                                list.Add(line);

                            list.Reverse();
                            var target = list.FirstOrDefault(l => l.Contains(index_model));
                            if (target == null)
                                return false;

                            State.opecode = target.Split(',')[2];

                            if (State.Category == State.CATEGORY.Q890)//Q890なら誤配線チェック合格していることを確認する
                                return target.Split(',')[6] == "00";
                            else
                                return true;//その他はログがあるだけでOK
                        }
                    };

                    //前段のログデータファイルを開き、読み込んだシリアルがあるかチェックする
                    if (State.Category == State.CATEGORY.Q890)
                    {
                        if (!System.IO.File.Exists(State.pc7LogFilePath) || !CheckLogFile(State.pc7LogFilePath))
                        {
                            MessageBox.Show("誤配線チェックが未実施です！");
                            timerTextInput.Start();
                            シリアル = false;
                            return;
                        }
                    }
                    else
                    {
                        if (!System.IO.File.Exists(State.pc3LogFilePath) || !CheckLogFile(State.pc3LogFilePath))
                        {
                            MessageBox.Show("配線ラベルが印刷されていません");
                            timerTextInput.Start();
                            シリアル = false;
                            return;
                        }
                    }


                    //PC4のログデータファイルを開き、読み込んだシリアルがあるかチェックする
                    if (System.IO.File.Exists(State.pc4LogFilePath) && CheckLogFile(State.pc4LogFilePath))
                    {
                        var result = MessageBox.Show("この製品は一度試験しています！\r\n再試験しますか？？", "警告", MessageBoxButtons.YesNo);
                        if (result == System.Windows.Forms.DialogResult.No)
                        {
                            timerTextInput.Start();
                            シリアル = false;
                            return;
                        }
                    }

                    シリアル = true; //ここでシリアル確定！

                    //該当する治具のパイロットランプを点滅させる処理を追加
                    timerLed.Start();


                    
                    var tm = new GeneralTimer(15000);
                    tm.Start();
                    while (true)
                    {
                        if (tm.FlagTimeout)
                        {
                            timerTextInput.Start();
                            シリアル = false;
                            timerLed.Stop();
                            Method.SetLed(false);
                            return;
                        }


                        //製品が正しく治具にセットされているかどうかのチェック
                        Method.io.ReadInputData(EPX64R.PortName.P7);
                        var SetBcr = (byte)(Method.io.P7InputData & 0x01) == 0x00;
                        var SetQ890 = (byte)(Method.io.P7InputData & 0x02) == 0x00;
                        var SetAur = (byte)(Method.io.P7InputData & 0x10) == 0x00;

                        switch (State.Category)
                        {
                            case State.CATEGORY.BCR:
                                if (SetBcr && !SetQ890 && !SetAur)
                                    goto PASS;
                                break;
                            case State.CATEGORY.Q890:
                                if (!SetBcr && SetQ890 && !SetAur)
                                    goto PASS;
                                break;
                            case State.CATEGORY.NEW_AUR:
                                if (!SetBcr && !SetQ890 && SetAur)
                                    goto PASS;
                                break;
                        }

                        Application.DoEvents();
                    }

                PASS:

                    //TODO:
                    //該当する治具のパイロットランプを消灯させる処理を追加
                    timerLed.Stop();
                    Method.SetLed(false);

                    //フォームのロック
                    LockForm(true);

                    timerLbMessage.Stop();

                    pictureBoxWarning.Visible = true;

                    //オペレーターへ電圧印加の警告をする
                    labelMessage.BackColor = Color.Yellow;
                    labelMessage.Location = new Point(115, 15);
                    labelMessage.Size = new Size(865, 75);

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

                    Method.InterLock解除();//インターロック解除

                    TOS9200.SendCommand("STOP", false);

                    //ｱｸﾉﾘｯｼﾞﾒｯｾｰｼﾞを返す設定
                    if (!TOS9200.SendCommand("SIL 0"))
                    {
                        MessageBox.Show("ｱｸﾉﾘｯｼﾞﾒｯｾｰｼﾞ設定異常");
                        timerTextInput.Start();
                        シリアル = false;

                        pictureBoxWarning.Visible = false;
                        labelMessage.Location = new Point(10, 15);
                        labelMessage.Size = new Size(965, 75);
                        labelMessage.BackColor = SystemColors.GradientActiveCaption;

                        timerLbMessage.Start();
                        ClearForm();
                        return;
                    }
                    MainTest();
                }
            }
            catch (Exception errorMessage)
            {
                MessageBox.Show("QRコード読み取りエラー\r\n" + errorMessage.Message, "警告！");
                Environment.Exit(0);
            }

        }

        private void textBoxSerial_TextChanged(object sender, EventArgs e)
        {
            //１文字入力されるごとに、タイマーを初期化する
            timerTextInput.Stop();
            timerTextInput.Start();
        }

        private void textBoxSerial_Enter(object sender, EventArgs e)
        {
            if (!作業者名 || !Flags.日常点検)
            {
                SetFocus();
            }
        }

        //タイマーイベント
        private void timerCurrentTime_Tick(object sender, EventArgs e)
        {

            //タイトルバーの更新
            this.Text = "BC-R 耐電圧/絶縁抵抗試験" + " <Ver_" + Constants.CheckerSoftVer + ">    " + "[" + System.DateTime.Now.ToString("yyyy年MM月dd日(ddd) hh時mm分ss秒") + "]";

        }

        private void timerLbMessage_Tick(object sender, EventArgs e)
        {
            if (labelMessage.BackColor == Color.Transparent)
            {
                labelMessage.BackColor = SystemColors.GradientActiveCaption;

                if (!作業者名)
                {
                    panelOperator.BackColor = SystemColors.GradientActiveCaption;
                    MenuItemUtility.BackColor = Color.Transparent;
                    panelSerial.BackColor = Color.Transparent;
                }
                else if (!Flags.日常点検)
                {
                    panelOperator.BackColor = Color.Transparent;
                    MenuItemUtility.BackColor = SystemColors.GradientActiveCaption;
                    panelSerial.BackColor = Color.Transparent;
                }
                else if (!シリアル)
                {
                    panelOperator.BackColor = Color.Transparent;
                    MenuItemUtility.BackColor = Color.Transparent;
                    panelSerial.BackColor = SystemColors.GradientActiveCaption;
                }
                else
                {
                    panelOperator.BackColor = Color.Transparent;
                    MenuItemUtility.BackColor = Color.Transparent;
                    panelSerial.BackColor = Color.Transparent;
                }


            }
            else
            {
                labelMessage.BackColor = Color.Transparent;
                panelOperator.BackColor = Color.Transparent;
                MenuItemUtility.BackColor = Color.Transparent;
                panelSerial.BackColor = Color.Transparent;

            }

        }

        private void timerCount_Tick(object sender, EventArgs e)//高電圧印加までの時間をカウントダウンする処理
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

        private void timerTextInput_Tick(object sender, EventArgs e)
        {
            timerTextInput.Stop();
            if (!シリアル)
            {
                textBoxSerial.Text = "";
            }

        }

        //コンボボックスイベント
        private void comboBoxOperatorName_DropDownClosed(object sender, EventArgs e)
        {
            if (comboBoxOperatorName.SelectedIndex == -1)
            {
                作業者名 = false;
            }
            else
            {
                作業者名 = true;
            }
        }

        private void comboBoxOperatorName_Leave(object sender, EventArgs e)
        {
            panelOperator.BackColor = Color.WhiteSmoke;
        }




        //フォーカスのセット**********************************************************************************
        private void SetFocus()
        {
            Application.DoEvents();
            if (!作業者名)
            {
                comboBoxOperatorName.Focus();
                labelMessage.Text = Constants.MessOperator;
                //General.PlaySound2(Constants.SoundName);
                return;
            }

            Application.DoEvents();
            if (!Flags.日常点検)
            {
                labelDailyCheck.Focus();
                labelMessage.Text = Constants.MessDailyCheck;
                //General.PlaySound2(Constants.SoundDailyCheck);
                return;
            }


            Application.DoEvents();
            if (!シリアル)
            {
                textBoxSerial.Focus();
                labelMessage.Text = Constants.MessQr;
                //General.PlaySound2(Constants.SoundSet);
                return;
            }

            Application.DoEvents();
            if (!SetJigu)
            {
                textBoxSerial.Focus();
                labelMessage.Text = Constants.MessSet;
                //General.PlaySound2(Constants.SoundSet);
                return;
            }

        }


        //プロパティ**********************************************************************************
        private bool 作業者名
        {
            get
            {
                return Flags.FlagOpeName;
            }
            set
            {
                State.OperatorName = value ? comboBoxOperatorName.SelectedItem.ToString() : "";
                Flags.FlagOpeName = value;
                this.comboBoxOperatorName.Enabled = !value;
                SetFocus();
            }
        }

        private bool シリアル
        {
            get
            {
                return Flags.FlagSerial;
            }
            set
            {
                Flags.FlagSerial = value;
                if (value)
                {
                    panelSerial.BackColor = Color.Transparent;
                }
                else
                {
                    textBoxSerial.Text = "";
                }
                //シリアルナンバーテキストボックスのロック
                textBoxSerial.Enabled = !value;

                SetFocus();
            }
        }

        private bool SetJigu
        {
            get
            {
                return Flags.FlagSetJigu;
            }
            set
            {
                Flags.FlagSetJigu = value;
                SetFocus();
            }
        }


        private void buttonClearOperator_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Tab:
                case Keys.Enter:
                    e.IsInputKey = true;
                    break;
                default: break;
            }
        }



        //**************************************************************************
        //フォームの初期化
        //引数：なし
        //戻値：なし
        //**************************************************************************
        private void InitForm()
        {

            //ラベルの初期化

            labelDanger.BackColor = Color.MistyRose;

            labelDecision.Text = "";
            labelErrorMessage.Text = "";

            pictureBoxWarning.Visible = false;
            labelMessage.Location = new Point(10, 15);
            labelMessage.Size = new Size(965, 75);

            labelMessage.Text = Constants.MessOperator;

            //BCR側の設定********************************************************************************************************************************
            //耐圧スペック
            labelBcr耐圧1_CH1.Text = Method.SetContact(State.耐電圧スペックリストBcr[0].CH1);
            labelBcr耐圧1_CH2.Text = Method.SetContact(State.耐電圧スペックリストBcr[0].CH2);
            labelBcr耐圧1_CH3.Text = Method.SetContact(State.耐電圧スペックリストBcr[0].CH3);
            labelBcr耐圧1_CH4.Text = Method.SetContact(State.耐電圧スペックリストBcr[0].CH4);
            labelBcr耐圧1_CH5.Text = Method.SetContact(State.耐電圧スペックリストBcr[0].CH5);
            labelBcr耐圧1_CH6.Text = Method.SetContact(State.耐電圧スペックリストBcr[0].CH6);
            labelBcr耐圧1Volt.Text = "AC" + State.耐電圧スペックリストBcr[0].印加電圧.ToString("F0") + "V";
            labelBcr耐圧1Time.Text = State.耐電圧スペックリストBcr[0].印加時間.ToString("F1") + "sec";
            labelBcr耐圧1Amp.Text = State.耐電圧スペックリストBcr[0].漏れ電流.ToString("F0") + "mA以下";
            labelBcr耐圧1計測値.Text = "---";

            labelBcr耐圧2_CH1.Text = Method.SetContact(State.耐電圧スペックリストBcr[1].CH1);
            labelBcr耐圧2_CH2.Text = Method.SetContact(State.耐電圧スペックリストBcr[1].CH2);
            labelBcr耐圧2_CH3.Text = Method.SetContact(State.耐電圧スペックリストBcr[1].CH3);
            labelBcr耐圧2_CH4.Text = Method.SetContact(State.耐電圧スペックリストBcr[1].CH4);
            labelBcr耐圧2_CH5.Text = Method.SetContact(State.耐電圧スペックリストBcr[1].CH5);
            labelBcr耐圧2_CH6.Text = Method.SetContact(State.耐電圧スペックリストBcr[1].CH6);
            labelBcr耐圧2Volt.Text = "AC" + State.耐電圧スペックリストBcr[1].印加電圧.ToString("F0") + "V";
            labelBcr耐圧2Time.Text = State.耐電圧スペックリストBcr[1].印加時間.ToString("F1") + "sec";
            labelBcr耐圧2Amp.Text = State.耐電圧スペックリストBcr[1].漏れ電流.ToString("F0") + "mA以下";
            labelBcr耐圧2計測値.Text = "---";


            //絶縁スペック
            labelBcr絶縁1_CH1.Text = Method.SetContact(State.絶縁抵抗スペックリストBcr[0].CH1);
            labelBcr絶縁1_CH2.Text = Method.SetContact(State.絶縁抵抗スペックリストBcr[0].CH2);
            labelBcr絶縁1_CH3.Text = Method.SetContact(State.絶縁抵抗スペックリストBcr[0].CH3);
            labelBcr絶縁1_CH4.Text = Method.SetContact(State.絶縁抵抗スペックリストBcr[0].CH4);
            labelBcr絶縁1_CH5.Text = Method.SetContact(State.絶縁抵抗スペックリストBcr[0].CH5);
            labelBcr絶縁1_CH6.Text = Method.SetContact(State.絶縁抵抗スペックリストBcr[0].CH6);
            labelBcr絶縁1Volt.Text = "DC" + State.絶縁抵抗スペックリストBcr[0].印加電圧.ToString("F0") + "V";
            labelBcr絶縁1Time.Text = State.絶縁抵抗スペックリストBcr[0].印加時間.ToString("F1") + "sec";
            labelBcr絶縁1Res.Text = State.絶縁抵抗スペックリストBcr[0].絶縁抵抗値.ToString("F0") + "MΩ以上";
            labelBcr絶縁1計測値.Text = "---";

            labelBcr絶縁2_CH1.Text = Method.SetContact(State.絶縁抵抗スペックリストBcr[1].CH1);
            labelBcr絶縁2_CH2.Text = Method.SetContact(State.絶縁抵抗スペックリストBcr[1].CH2);
            labelBcr絶縁2_CH3.Text = Method.SetContact(State.絶縁抵抗スペックリストBcr[1].CH3);
            labelBcr絶縁2_CH4.Text = Method.SetContact(State.絶縁抵抗スペックリストBcr[1].CH4);
            labelBcr絶縁2_CH5.Text = Method.SetContact(State.絶縁抵抗スペックリストBcr[1].CH5);
            labelBcr絶縁2_CH6.Text = Method.SetContact(State.絶縁抵抗スペックリストBcr[1].CH6);
            labelBcr絶縁2Volt.Text = "DC" + State.絶縁抵抗スペックリストBcr[1].印加電圧.ToString("F0") + "V";
            labelBcr絶縁2Time.Text = State.絶縁抵抗スペックリストBcr[1].印加時間.ToString("F1") + "sec";
            labelBcr絶縁2Res.Text = State.絶縁抵抗スペックリストBcr[1].絶縁抵抗値.ToString("F0") + "MΩ以上";
            labelBcr絶縁2計測値.Text = "---";

            labelBcr絶縁3_CH1.Text = Method.SetContact(State.絶縁抵抗スペックリストBcr[2].CH1);
            labelBcr絶縁3_CH2.Text = Method.SetContact(State.絶縁抵抗スペックリストBcr[2].CH2);
            labelBcr絶縁3_CH3.Text = Method.SetContact(State.絶縁抵抗スペックリストBcr[2].CH3);
            labelBcr絶縁3_CH4.Text = Method.SetContact(State.絶縁抵抗スペックリストBcr[2].CH4);
            labelBcr絶縁3_CH5.Text = Method.SetContact(State.絶縁抵抗スペックリストBcr[2].CH5);
            labelBcr絶縁3_CH6.Text = Method.SetContact(State.絶縁抵抗スペックリストBcr[2].CH6);
            labelBcr絶縁3Volt.Text = "DC" + State.絶縁抵抗スペックリストBcr[2].印加電圧.ToString("F0") + "V";
            labelBcr絶縁3Time.Text = State.絶縁抵抗スペックリストBcr[2].印加時間.ToString("F1") + "sec";
            labelBcr絶縁3Res.Text = State.絶縁抵抗スペックリストBcr[2].絶縁抵抗値.ToString("F0") + "MΩ以上";
            labelBcr絶縁3計測値.Text = "---";

            labelBcr耐圧1計測値.ForeColor = Color.Black;
            labelBcr耐圧2計測値.ForeColor = Color.Black;
            labelBcr絶縁1計測値.ForeColor = Color.Black;
            labelBcr絶縁2計測値.ForeColor = Color.Black;
            labelBcr絶縁3計測値.ForeColor = Color.Black;


            //Q890側の設定************************************************************************************************************************************

            //耐圧スペック
            labelQ890耐圧1_CH1.Text = Method.SetContact(State.耐電圧スペックリストQ890[0].CH1);
            labelQ890耐圧1_CH2.Text = Method.SetContact(State.耐電圧スペックリストQ890[0].CH2);
            labelQ890耐圧1_CH3.Text = Method.SetContact(State.耐電圧スペックリストQ890[0].CH3);
            labelQ890耐圧1_CH4.Text = Method.SetContact(State.耐電圧スペックリストQ890[0].CH4);
            labelQ890耐圧1Volt.Text = "AC" + State.耐電圧スペックリストQ890[0].印加電圧.ToString("F0") + "V";
            labelQ890耐圧1Time.Text = State.耐電圧スペックリストQ890[0].印加時間.ToString("F1") + "sec";
            labelQ890耐圧1Amp.Text = State.耐電圧スペックリストQ890[0].漏れ電流.ToString("F0") + "mA以下";
            labelQ890耐圧1計測値.Text = "---";


            //絶縁スペック
            labelQ890絶縁1_CH1.Text = Method.SetContact(State.絶縁抵抗スペックリストQ890[0].CH1);
            labelQ890絶縁1_CH2.Text = Method.SetContact(State.絶縁抵抗スペックリストQ890[0].CH2);
            labelQ890絶縁1_CH3.Text = Method.SetContact(State.絶縁抵抗スペックリストQ890[0].CH3);
            labelQ890絶縁1_CH4.Text = Method.SetContact(State.絶縁抵抗スペックリストQ890[0].CH4);
            labelQ890絶縁1Volt.Text = "DC" + State.絶縁抵抗スペックリストQ890[0].印加電圧.ToString("F0") + "V";
            labelQ890絶縁1Time.Text = State.絶縁抵抗スペックリストQ890[0].印加時間.ToString("F1") + "sec";
            labelQ890絶縁1Res.Text = State.絶縁抵抗スペックリストQ890[0].絶縁抵抗値.ToString("F0") + "MΩ以上";
            labelQ890絶縁1計測値.Text = "---";

            labelQ890耐圧1計測値.ForeColor = Color.Black;
            labelQ890絶縁1計測値.ForeColor = Color.Black;


            //AUR側の設定******************************************************************************************************************************

            //耐圧スペック
            labelNewAur耐圧1_CH1.Text = Method.SetContact(State.耐電圧スペックリストAur[0].CH1);
            labelNewAur耐圧1_CH2.Text = Method.SetContact(State.耐電圧スペックリストAur[0].CH2);
            labelNewAur耐圧1_CH3.Text = Method.SetContact(State.耐電圧スペックリストAur[0].CH3);
            labelNewAur耐圧1_CH4.Text = Method.SetContact(State.耐電圧スペックリストAur[0].CH4);
            labelNewAur耐圧1_CH5.Text = Method.SetContact(State.耐電圧スペックリストAur[0].CH5);
            labelNewAur耐圧1_CH6.Text = Method.SetContact(State.耐電圧スペックリストAur[0].CH6);
            labelNewAur耐圧1Volt.Text = "AC" + State.耐電圧スペックリストAur[0].印加電圧.ToString("F0") + "V";
            labelNewAur耐圧1Time.Text = State.耐電圧スペックリストAur[0].印加時間.ToString("F1") + "sec";
            labelNewAur耐圧1Amp.Text = State.耐電圧スペックリストAur[0].漏れ電流.ToString("F0") + "mA以下";
            labelNewAur耐圧1計測値.Text = "---";

            labelNewAur耐圧2_CH1.Text = Method.SetContact(State.耐電圧スペックリストAur[1].CH1);
            labelNewAur耐圧2_CH2.Text = Method.SetContact(State.耐電圧スペックリストAur[1].CH2);
            labelNewAur耐圧2_CH3.Text = Method.SetContact(State.耐電圧スペックリストAur[1].CH3);
            labelNewAur耐圧2_CH4.Text = Method.SetContact(State.耐電圧スペックリストAur[1].CH4);
            labelNewAur耐圧2_CH5.Text = Method.SetContact(State.耐電圧スペックリストAur[1].CH5);
            labelNewAur耐圧2_CH6.Text = Method.SetContact(State.耐電圧スペックリストAur[1].CH6);
            labelNewAur耐圧2Volt.Text = "AC" + State.耐電圧スペックリストAur[1].印加電圧.ToString("F0") + "V";
            labelNewAur耐圧2Time.Text = State.耐電圧スペックリストAur[1].印加時間.ToString("F1") + "sec";
            labelNewAur耐圧2Amp.Text = State.耐電圧スペックリストAur[1].漏れ電流.ToString("F0") + "mA以下";
            labelNewAur耐圧2計測値.Text = "---";

            labelNewAur耐圧3_CH1.Text = Method.SetContact(State.耐電圧スペックリストAur[2].CH1);
            labelNewAur耐圧3_CH2.Text = Method.SetContact(State.耐電圧スペックリストAur[2].CH2);
            labelNewAur耐圧3_CH3.Text = Method.SetContact(State.耐電圧スペックリストAur[2].CH3);
            labelNewAur耐圧3_CH4.Text = Method.SetContact(State.耐電圧スペックリストAur[2].CH4);
            labelNewAur耐圧3_CH5.Text = Method.SetContact(State.耐電圧スペックリストAur[2].CH5);
            labelNewAur耐圧3_CH6.Text = Method.SetContact(State.耐電圧スペックリストAur[2].CH6);
            labelNewAur耐圧3Volt.Text = "AC" + State.耐電圧スペックリストAur[2].印加電圧.ToString("F0") + "V";
            labelNewAur耐圧3Time.Text = State.耐電圧スペックリストAur[2].印加時間.ToString("F1") + "sec";
            labelNewAur耐圧3Amp.Text = State.耐電圧スペックリストAur[2].漏れ電流.ToString("F0") + "mA以下";
            labelNewAur耐圧3計測値.Text = "---";


            //絶縁スペック
            labelNewAur絶縁1_CH1.Text = Method.SetContact(State.絶縁抵抗スペックリストAur[0].CH1);
            labelNewAur絶縁1_CH2.Text = Method.SetContact(State.絶縁抵抗スペックリストAur[0].CH2);
            labelNewAur絶縁1_CH3.Text = Method.SetContact(State.絶縁抵抗スペックリストAur[0].CH3);
            labelNewAur絶縁1_CH4.Text = Method.SetContact(State.絶縁抵抗スペックリストAur[0].CH4);
            labelNewAur絶縁1_CH5.Text = Method.SetContact(State.絶縁抵抗スペックリストAur[0].CH5);
            labelNewAur絶縁1_CH6.Text = Method.SetContact(State.絶縁抵抗スペックリストAur[0].CH6);
            labelNewAur絶縁1Volt.Text = "DC" + State.絶縁抵抗スペックリストAur[0].印加電圧.ToString("F0") + "V";
            labelNewAur絶縁1Time.Text = State.絶縁抵抗スペックリストAur[0].印加時間.ToString("F1") + "sec";
            labelNewAur絶縁1Res.Text = State.絶縁抵抗スペックリストAur[0].絶縁抵抗値.ToString("F0") + "MΩ以上";
            labelNewAur絶縁1計測値.Text = "---";

            labelNewAur絶縁2_CH1.Text = Method.SetContact(State.絶縁抵抗スペックリストAur[1].CH1);
            labelNewAur絶縁2_CH2.Text = Method.SetContact(State.絶縁抵抗スペックリストAur[1].CH2);
            labelNewAur絶縁2_CH3.Text = Method.SetContact(State.絶縁抵抗スペックリストAur[1].CH3);
            labelNewAur絶縁2_CH4.Text = Method.SetContact(State.絶縁抵抗スペックリストAur[1].CH4);
            labelNewAur絶縁2_CH5.Text = Method.SetContact(State.絶縁抵抗スペックリストAur[1].CH5);
            labelNewAur絶縁2_CH6.Text = Method.SetContact(State.絶縁抵抗スペックリストAur[1].CH6);
            labelNewAur絶縁2Volt.Text = "DC" + State.絶縁抵抗スペックリストAur[1].印加電圧.ToString("F0") + "V";
            labelNewAur絶縁2Time.Text = State.絶縁抵抗スペックリストAur[1].印加時間.ToString("F1") + "sec";
            labelNewAur絶縁2Res.Text = State.絶縁抵抗スペックリストAur[1].絶縁抵抗値.ToString("F0") + "MΩ以上";
            labelNewAur絶縁2計測値.Text = "---";

            labelNewAur絶縁3_CH1.Text = Method.SetContact(State.絶縁抵抗スペックリストAur[2].CH1);
            labelNewAur絶縁3_CH2.Text = Method.SetContact(State.絶縁抵抗スペックリストAur[2].CH2);
            labelNewAur絶縁3_CH3.Text = Method.SetContact(State.絶縁抵抗スペックリストAur[2].CH3);
            labelNewAur絶縁3_CH4.Text = Method.SetContact(State.絶縁抵抗スペックリストAur[2].CH4);
            labelNewAur絶縁3_CH5.Text = Method.SetContact(State.絶縁抵抗スペックリストAur[2].CH5);
            labelNewAur絶縁3_CH6.Text = Method.SetContact(State.絶縁抵抗スペックリストAur[2].CH6);
            labelNewAur絶縁3Volt.Text = "DC" + State.絶縁抵抗スペックリストAur[2].印加電圧.ToString("F0") + "V";
            labelNewAur絶縁3Time.Text = State.絶縁抵抗スペックリストAur[2].印加時間.ToString("F1") + "sec";
            labelNewAur絶縁3Res.Text = State.絶縁抵抗スペックリストAur[2].絶縁抵抗値.ToString("F0") + "MΩ以上";
            labelNewAur絶縁3計測値.Text = "---";

            labelNewAur耐圧1計測値.ForeColor = Color.Black;
            labelNewAur耐圧2計測値.ForeColor = Color.Black;
            labelNewAur耐圧3計測値.ForeColor = Color.Black;

            labelNewAur絶縁1計測値.ForeColor = Color.Black;
            labelNewAur絶縁2計測値.ForeColor = Color.Black;
            labelNewAur絶縁3計測値.ForeColor = Color.Black;


            //テキストボックスの初期化
            textBoxSerial.Text = "";
            textBoxSerial.Enabled = true;


            //作業者一覧の設定
            comboBoxOperatorName.Items.Clear();
            foreach (string Operator in State.作業者リスト)
            {
                comboBoxOperatorName.Items.Add(Operator);
            }
            comboBoxOperatorName.Text = "";



            作業者名 = false;
            シリアル = false;

            SetFocus();

        }

        //**************************************************************************
        //試験終了後のフォームのクリア
        //引数：なし
        //戻値：なし
        //**************************************************************************
        private void ClearForm()
        {
            //ラベルのクリア
            labelDecision.Text = "";
            labelErrorMessage.Text = "";

            labelBcr耐圧1計測値.Text = "---";
            labelBcr耐圧2計測値.Text = "---";
            labelBcr絶縁1計測値.Text = "---";
            labelBcr絶縁2計測値.Text = "---";
            labelBcr絶縁3計測値.Text = "---";

            labelQ890耐圧1計測値.Text = "---";
            labelQ890絶縁1計測値.Text = "---";

            labelNewAur耐圧1計測値.Text = "---";
            labelNewAur耐圧2計測値.Text = "---";
            labelNewAur耐圧3計測値.Text = "---";
            labelNewAur絶縁1計測値.Text = "---";
            labelNewAur絶縁2計測値.Text = "---";
            labelNewAur絶縁3計測値.Text = "---";


            labelBcr耐圧1計測値.ForeColor = Color.Black;
            labelBcr耐圧2計測値.ForeColor = Color.Black;
            labelBcr絶縁1計測値.ForeColor = Color.Black;
            labelBcr絶縁2計測値.ForeColor = Color.Black;
            labelBcr絶縁3計測値.ForeColor = Color.Black;

            labelQ890耐圧1計測値.ForeColor = Color.Black;
            labelQ890絶縁1計測値.ForeColor = Color.Black;

            labelNewAur耐圧1計測値.ForeColor = Color.Black;
            labelNewAur耐圧2計測値.ForeColor = Color.Black;
            labelNewAur耐圧3計測値.ForeColor = Color.Black;
            labelNewAur絶縁1計測値.ForeColor = Color.Black;
            labelNewAur絶縁2計測値.ForeColor = Color.Black;
            labelNewAur絶縁3計測値.ForeColor = Color.Black;

            LockForm(false);

            //テキストボックスのクリア
            シリアル = false;
            textBoxSerial.Focus();



        }

        //**************************************************************************
        //フォームのロック
        //引数：bool
        //戻値：true：ロック　false：ロック解除
        //**************************************************************************
        private void LockForm(bool data)
        {
            if (data)
            {
                // 最小化ボタンを無効にする
                this.MinimizeBox = false;
                textBoxSerial.Enabled = false;
                menuStrip1.Enabled = false;
                buttonClearOperator.Enabled = false;
            }
            else
            {
                // 最小化ボタンを無効にする
                this.MinimizeBox = true;
                textBoxSerial.Enabled = true;
                menuStrip1.Enabled = true;
                buttonClearOperator.Enabled = true;
            }

        }


        //**************************************************************************
        //試験スペック表示ラベルの色を変える
        //引数：なし
        //戻値：なし
        //**************************************************************************
        private void SetSpecLabelColor(耐電圧試験スペック spec = null)
        {
            if (spec == null)
            {
                if (State.Category == State.CATEGORY.BCR)
                {

                    labelBcr耐圧1ステップ.BackColor = Color.Transparent;
                    labelBcr耐圧1_CH1.BackColor = Color.Transparent;
                    labelBcr耐圧1_CH2.BackColor = Color.Transparent;
                    labelBcr耐圧1_CH3.BackColor = Color.Transparent;
                    labelBcr耐圧1_CH4.BackColor = Color.Transparent;
                    labelBcr耐圧1_CH5.BackColor = Color.Transparent;
                    labelBcr耐圧1_CH6.BackColor = Color.Transparent;
                    labelBcr耐圧1Volt.BackColor = Color.Transparent;
                    labelBcr耐圧1Time.BackColor = Color.Transparent;
                    labelBcr耐圧1Amp.BackColor = Color.Transparent;

                    labelBcr耐圧2ステップ.BackColor = Color.Transparent;
                    labelBcr耐圧2_CH1.BackColor = Color.Transparent;
                    labelBcr耐圧2_CH2.BackColor = Color.Transparent;
                    labelBcr耐圧2_CH3.BackColor = Color.Transparent;
                    labelBcr耐圧2_CH4.BackColor = Color.Transparent;
                    labelBcr耐圧2_CH5.BackColor = Color.Transparent;
                    labelBcr耐圧2_CH6.BackColor = Color.Transparent;
                    labelBcr耐圧2Volt.BackColor = Color.Transparent;
                    labelBcr耐圧2Time.BackColor = Color.Transparent;
                    labelBcr耐圧2Amp.BackColor = Color.Transparent;

                }
                else if (State.Category == State.CATEGORY.Q890)
                {
                    labelQ890耐圧1ステップ.BackColor = Color.Transparent;
                    labelQ890耐圧1_CH1.BackColor = Color.Transparent;
                    labelQ890耐圧1_CH2.BackColor = Color.Transparent;
                    labelQ890耐圧1_CH3.BackColor = Color.Transparent;
                    labelQ890耐圧1_CH4.BackColor = Color.Transparent;
                    labelQ890耐圧1Volt.BackColor = Color.Transparent;
                    labelQ890耐圧1Time.BackColor = Color.Transparent;
                    labelQ890耐圧1Amp.BackColor = Color.Transparent;

                }
                else
                {
                    labelNewAur耐圧1ステップ.BackColor = Color.Transparent;
                    labelNewAur耐圧1_CH1.BackColor = Color.Transparent;
                    labelNewAur耐圧1_CH2.BackColor = Color.Transparent;
                    labelNewAur耐圧1_CH3.BackColor = Color.Transparent;
                    labelNewAur耐圧1_CH4.BackColor = Color.Transparent;
                    labelNewAur耐圧1_CH5.BackColor = Color.Transparent;
                    labelNewAur耐圧1_CH6.BackColor = Color.Transparent;
                    labelNewAur耐圧1Volt.BackColor = Color.Transparent;
                    labelNewAur耐圧1Time.BackColor = Color.Transparent;
                    labelNewAur耐圧1Amp.BackColor = Color.Transparent;

                    labelNewAur耐圧2ステップ.BackColor = Color.Transparent;
                    labelNewAur耐圧2_CH1.BackColor = Color.Transparent;
                    labelNewAur耐圧2_CH2.BackColor = Color.Transparent;
                    labelNewAur耐圧2_CH3.BackColor = Color.Transparent;
                    labelNewAur耐圧2_CH4.BackColor = Color.Transparent;
                    labelNewAur耐圧2_CH5.BackColor = Color.Transparent;
                    labelNewAur耐圧2_CH6.BackColor = Color.Transparent;
                    labelNewAur耐圧2Volt.BackColor = Color.Transparent;
                    labelNewAur耐圧2Time.BackColor = Color.Transparent;
                    labelNewAur耐圧2Amp.BackColor = Color.Transparent;

                    labelNewAur耐圧3ステップ.BackColor = Color.Transparent;
                    labelNewAur耐圧3_CH1.BackColor = Color.Transparent;
                    labelNewAur耐圧3_CH2.BackColor = Color.Transparent;
                    labelNewAur耐圧3_CH3.BackColor = Color.Transparent;
                    labelNewAur耐圧3_CH4.BackColor = Color.Transparent;
                    labelNewAur耐圧3_CH5.BackColor = Color.Transparent;
                    labelNewAur耐圧3_CH6.BackColor = Color.Transparent;
                    labelNewAur耐圧3Volt.BackColor = Color.Transparent;
                    labelNewAur耐圧3Time.BackColor = Color.Transparent;
                    labelNewAur耐圧3Amp.BackColor = Color.Transparent;

                }
                return;

            }

            if (State.Category == State.CATEGORY.BCR)
            {
                switch (spec.ステップ)
                {

                    case "STEP1":
                        labelBcr耐圧1ステップ.BackColor = Color.LightPink;
                        labelBcr耐圧1_CH1.BackColor = Color.LightPink;
                        labelBcr耐圧1_CH2.BackColor = Color.LightPink;
                        labelBcr耐圧1_CH3.BackColor = Color.LightPink;
                        labelBcr耐圧1_CH4.BackColor = Color.LightPink;
                        labelBcr耐圧1_CH5.BackColor = Color.LightPink;
                        labelBcr耐圧1_CH6.BackColor = Color.LightPink;
                        labelBcr耐圧1Volt.BackColor = Color.LightPink;
                        labelBcr耐圧1Time.BackColor = Color.LightPink;
                        labelBcr耐圧1Amp.BackColor = Color.LightPink;
                        break;

                    case "STEP2":
                        labelBcr耐圧2ステップ.BackColor = Color.LightPink;
                        labelBcr耐圧2_CH1.BackColor = Color.LightPink;
                        labelBcr耐圧2_CH2.BackColor = Color.LightPink;
                        labelBcr耐圧2_CH3.BackColor = Color.LightPink;
                        labelBcr耐圧2_CH4.BackColor = Color.LightPink;
                        labelBcr耐圧2_CH5.BackColor = Color.LightPink;
                        labelBcr耐圧2_CH6.BackColor = Color.LightPink;
                        labelBcr耐圧2Volt.BackColor = Color.LightPink;
                        labelBcr耐圧2Time.BackColor = Color.LightPink;
                        labelBcr耐圧2Amp.BackColor = Color.LightPink;
                        break;

                    default: break;
                }
            }
            else if (State.Category == State.CATEGORY.Q890)
            {
                switch (spec.ステップ)
                {
                    case "STEP1":
                        labelQ890耐圧1ステップ.BackColor = Color.LightPink;
                        labelQ890耐圧1_CH1.BackColor = Color.LightPink;
                        labelQ890耐圧1_CH2.BackColor = Color.LightPink;
                        labelQ890耐圧1_CH3.BackColor = Color.LightPink;
                        labelQ890耐圧1_CH4.BackColor = Color.LightPink;
                        labelQ890耐圧1Volt.BackColor = Color.LightPink;
                        labelQ890耐圧1Time.BackColor = Color.LightPink;
                        labelQ890耐圧1Amp.BackColor = Color.LightPink;
                        break;

                    default: break;
                }
            }
            else
            {
                switch (spec.ステップ)
                {

                    case "STEP1":
                        labelNewAur耐圧1ステップ.BackColor = Color.LightPink;
                        labelNewAur耐圧1_CH1.BackColor = Color.LightPink;
                        labelNewAur耐圧1_CH2.BackColor = Color.LightPink;
                        labelNewAur耐圧1_CH3.BackColor = Color.LightPink;
                        labelNewAur耐圧1_CH4.BackColor = Color.LightPink;
                        labelNewAur耐圧1_CH5.BackColor = Color.LightPink;
                        labelNewAur耐圧1_CH6.BackColor = Color.LightPink;
                        labelNewAur耐圧1Volt.BackColor = Color.LightPink;
                        labelNewAur耐圧1Time.BackColor = Color.LightPink;
                        labelNewAur耐圧1Amp.BackColor = Color.LightPink;
                        break;

                    case "STEP2":
                        labelNewAur耐圧2ステップ.BackColor = Color.LightPink;
                        labelNewAur耐圧2_CH1.BackColor = Color.LightPink;
                        labelNewAur耐圧2_CH2.BackColor = Color.LightPink;
                        labelNewAur耐圧2_CH3.BackColor = Color.LightPink;
                        labelNewAur耐圧2_CH4.BackColor = Color.LightPink;
                        labelNewAur耐圧2_CH5.BackColor = Color.LightPink;
                        labelNewAur耐圧2_CH6.BackColor = Color.LightPink;
                        labelNewAur耐圧2Volt.BackColor = Color.LightPink;
                        labelNewAur耐圧2Time.BackColor = Color.LightPink;
                        labelNewAur耐圧2Amp.BackColor = Color.LightPink;
                        break;

                    case "STEP3":
                        labelNewAur耐圧3ステップ.BackColor = Color.LightPink;
                        labelNewAur耐圧3_CH1.BackColor = Color.LightPink;
                        labelNewAur耐圧3_CH2.BackColor = Color.LightPink;
                        labelNewAur耐圧3_CH3.BackColor = Color.LightPink;
                        labelNewAur耐圧3_CH4.BackColor = Color.LightPink;
                        labelNewAur耐圧3_CH5.BackColor = Color.LightPink;
                        labelNewAur耐圧3_CH6.BackColor = Color.LightPink;
                        labelNewAur耐圧3Volt.BackColor = Color.LightPink;
                        labelNewAur耐圧3Time.BackColor = Color.LightPink;
                        labelNewAur耐圧3Amp.BackColor = Color.LightPink;
                        break;

                    default: break;
                }

            }
        }


        private void SetSpecLabelColor(絶縁抵抗試験スペック spec = null) //オーバーロード
        {
            Action<Color> Set絶縁1 = (color) =>
            {
                if (State.Category == State.CATEGORY.BCR)
                {
                    labelBcr絶縁1ステップ.BackColor = color;
                    labelBcr絶縁1_CH1.BackColor = color;
                    labelBcr絶縁1_CH2.BackColor = color;
                    labelBcr絶縁1_CH3.BackColor = color;
                    labelBcr絶縁1_CH4.BackColor = color;
                    labelBcr絶縁1_CH5.BackColor = color;
                    labelBcr絶縁1_CH6.BackColor = color;
                    labelBcr絶縁1Volt.BackColor = color;
                    labelBcr絶縁1Time.BackColor = color;
                    labelBcr絶縁1Res.BackColor = color;
                }
                else if (State.Category == State.CATEGORY.Q890)
                {
                    labelQ890絶縁1ステップ.BackColor = color;
                    labelQ890絶縁1_CH1.BackColor = color;
                    labelQ890絶縁1_CH2.BackColor = color;
                    labelQ890絶縁1_CH3.BackColor = color;
                    labelQ890絶縁1_CH4.BackColor = color;
                    labelQ890絶縁1Volt.BackColor = color;
                    labelQ890絶縁1Time.BackColor = color;
                    labelQ890絶縁1Res.BackColor = color;
                }
                else
                {
                    labelNewAur絶縁1ステップ.BackColor = color;
                    labelNewAur絶縁1_CH1.BackColor = color;
                    labelNewAur絶縁1_CH2.BackColor = color;
                    labelNewAur絶縁1_CH3.BackColor = color;
                    labelNewAur絶縁1_CH4.BackColor = color;
                    labelNewAur絶縁1_CH5.BackColor = color;
                    labelNewAur絶縁1_CH6.BackColor = color;
                    labelNewAur絶縁1Volt.BackColor = color;
                    labelNewAur絶縁1Time.BackColor = color;
                    labelNewAur絶縁1Res.BackColor = color;
                }

            };

            Action<Color> Set絶縁2 = (color) =>
            {
                if (State.Category == State.CATEGORY.BCR)
                {
                    labelBcr絶縁2ステップ.BackColor = color;
                    labelBcr絶縁2_CH1.BackColor = color;
                    labelBcr絶縁2_CH2.BackColor = color;
                    labelBcr絶縁2_CH3.BackColor = color;
                    labelBcr絶縁2_CH4.BackColor = color;
                    labelBcr絶縁2_CH5.BackColor = color;
                    labelBcr絶縁2_CH6.BackColor = color;
                    labelBcr絶縁2Volt.BackColor = color;
                    labelBcr絶縁2Time.BackColor = color;
                    labelBcr絶縁2Res.BackColor = color;
                }
                else
                {
                    labelNewAur絶縁2ステップ.BackColor = color;
                    labelNewAur絶縁2_CH1.BackColor = color;
                    labelNewAur絶縁2_CH2.BackColor = color;
                    labelNewAur絶縁2_CH3.BackColor = color;
                    labelNewAur絶縁2_CH4.BackColor = color;
                    labelNewAur絶縁2_CH5.BackColor = color;
                    labelNewAur絶縁2_CH6.BackColor = color;
                    labelNewAur絶縁2Volt.BackColor = color;
                    labelNewAur絶縁2Time.BackColor = color;
                    labelNewAur絶縁2Res.BackColor = color;
                }
            };

            Action<Color> Set絶縁3 = (color) =>
            {
                if (State.Category == State.CATEGORY.BCR)
                {
                    labelBcr絶縁3ステップ.BackColor = color;
                    labelBcr絶縁3_CH1.BackColor = color;
                    labelBcr絶縁3_CH2.BackColor = color;
                    labelBcr絶縁3_CH3.BackColor = color;
                    labelBcr絶縁3_CH4.BackColor = color;
                    labelBcr絶縁3_CH5.BackColor = color;
                    labelBcr絶縁3_CH6.BackColor = color;
                    labelBcr絶縁3Volt.BackColor = color;
                    labelBcr絶縁3Time.BackColor = color;
                    labelBcr絶縁3Res.BackColor = color;
                }
                else
                {
                    labelNewAur絶縁3ステップ.BackColor = color;
                    labelNewAur絶縁3_CH1.BackColor = color;
                    labelNewAur絶縁3_CH2.BackColor = color;
                    labelNewAur絶縁3_CH3.BackColor = color;
                    labelNewAur絶縁3_CH4.BackColor = color;
                    labelNewAur絶縁3_CH5.BackColor = color;
                    labelNewAur絶縁3_CH6.BackColor = color;
                    labelNewAur絶縁3Volt.BackColor = color;
                    labelNewAur絶縁3Time.BackColor = color;
                    labelNewAur絶縁3Res.BackColor = color;
                }
            };


            if (spec == null)
            {
                Set絶縁1(Color.Transparent);
                Set絶縁2(Color.Transparent);
                Set絶縁3(Color.Transparent);
                return;
            }

            switch (spec.ステップ)
            {
                case "STEP1":
                    Set絶縁1(Color.LightPink);
                    Set絶縁2(Color.Transparent);
                    Set絶縁3(Color.Transparent);
                    break;

                case "STEP2":
                    Set絶縁1(Color.Transparent);
                    Set絶縁2(Color.LightPink);
                    Set絶縁3(Color.Transparent);
                    break;

                case "STEP3":
                    Set絶縁1(Color.Transparent);
                    Set絶縁2(Color.Transparent);
                    Set絶縁3(Color.LightPink);
                    break;

                default: break;
            }

        }




        //メインルーチン（本体試験）
        public void MainTest()
        {
            Flags.FlagTest = true;


            //各オブジェクトの計測値データをクリア
            foreach (var spec in State.耐電圧スペックリストBcr)
            {
                spec.計測値 = "";
            }

            foreach (var spec in State.絶縁抵抗スペックリストBcr)
            {
                spec.計測値 = "";
            }

            //各オブジェクトの計測値データをクリア
            foreach (var spec in State.耐電圧スペックリストQ890)
            {
                spec.計測値 = "";
            }

            foreach (var spec in State.絶縁抵抗スペックリストQ890)
            {
                spec.計測値 = "";
            }


            labelMessage.BackColor = Color.Red;
            labelMessage.Text = Constants.MessWait;
            labelDanger.BackColor = Color.Red;

            TOS9200.SendCommand("*CLS");
            TOS9200.SendCommand("STOP");


            //①耐電圧試験**********************************************************************************************************************************
            if (State.Category == State.CATEGORY.BCR)
            {
                foreach (var spec in State.耐電圧スペックリストBcr)
                {
                    Application.DoEvents();

                    TOS9200.Dsr = 0;    //DSRレジスタ初期化
                    TOS9200.Fail = 0;   //failレジスタ初期化
                    if (!TOS9200.SetFunction(spec)) goto NgStep;//試験設定
                    SetSpecLabelColor((耐電圧試験スペック)null);//すべてのラベル色を戻すための処理
                    SetSpecLabelColor(spec);

                    TOS9200.TosStart();

                    spec.計測値 = TOS9200.GetMeasureData() + "A"; //計測データの取得

                    switch (spec.ステップ)
                    {
                        case "STEP1":
                            labelBcr耐圧1計測値.Text = spec.計測値;
                            if (TOS9200.ErrorMess != TOS9200.ErrorMessage.正常終了)
                            {
                                labelBcr耐圧1計測値.ForeColor = Color.Red;
                                goto NgStep;
                            }
                            break;

                        case "STEP2":
                            labelBcr耐圧2計測値.Text = spec.計測値;
                            if (TOS9200.ErrorMess != TOS9200.ErrorMessage.正常終了)
                            {
                                labelBcr耐圧2計測値.ForeColor = Color.Red;
                                goto NgStep;
                            }
                            break;

                        default: break;
                    }
                    TOS9200.SendCommand("STOP");
                }
            }
            else if (State.Category == State.CATEGORY.Q890)
            {
                foreach (var spec in State.耐電圧スペックリストQ890)
                {
                    Application.DoEvents();

                    TOS9200.Dsr = 0;    //DSRレジスタ初期化
                    TOS9200.Fail = 0;   //failレジスタ初期化
                    TOS9200.SetFunction(spec);//試験設定
                    SetSpecLabelColor(spec);

                    TOS9200.TosStart();

                    spec.計測値 = TOS9200.GetMeasureData() + "A"; //計測データの取得

                    switch (spec.ステップ)
                    {
                        case "STEP1":
                            labelQ890耐圧1計測値.Text = spec.計測値;
                            if (TOS9200.ErrorMess != TOS9200.ErrorMessage.正常終了)
                            {
                                labelQ890耐圧1計測値.ForeColor = Color.Red;
                                goto NgStep;
                            }
                            break;

                        default: break;
                    }
                    TOS9200.SendCommand("STOP");
                }
            }
            else
            {
                foreach (var spec in State.耐電圧スペックリストAur)
                {
                    Application.DoEvents();

                    TOS9200.Dsr = 0;    //DSRレジスタ初期化
                    TOS9200.Fail = 0;   //failレジスタ初期化
                    if (!TOS9200.SetFunction(spec)) goto NgStep;//試験設定
                    SetSpecLabelColor((耐電圧試験スペック)null);//すべてのラベル色を戻すための処理
                    SetSpecLabelColor(spec);

                    TOS9200.TosStart();

                    spec.計測値 = TOS9200.GetMeasureData() + "A"; //計測データの取得

                    switch (spec.ステップ)
                    {
                        case "STEP1":
                            labelNewAur耐圧1計測値.Text = spec.計測値;
                            if (TOS9200.ErrorMess != TOS9200.ErrorMessage.正常終了)
                            {
                                labelNewAur耐圧1計測値.ForeColor = Color.Red;
                                goto NgStep;
                            }
                            break;

                        case "STEP2":
                            labelNewAur耐圧2計測値.Text = spec.計測値;
                            if (TOS9200.ErrorMess != TOS9200.ErrorMessage.正常終了)
                            {
                                labelNewAur耐圧2計測値.ForeColor = Color.Red;
                                goto NgStep;
                            }
                            break;

                        case "STEP3":
                            labelNewAur耐圧3計測値.Text = spec.計測値;
                            if (TOS9200.ErrorMess != TOS9200.ErrorMessage.正常終了)
                            {
                                labelNewAur耐圧3計測値.ForeColor = Color.Red;
                                goto NgStep;
                            }
                            break;

                        default: break;
                    }
                    TOS9200.SendCommand("STOP");
                }

            }



            //②絶縁抵抗試験************************************************************************************************************************************
            SetSpecLabelColor((耐電圧試験スペック)null);//すべてのラベル色を戻すための処理
            if (State.Category == State.CATEGORY.BCR)
            {
                foreach (var spec in State.絶縁抵抗スペックリストBcr)
                {
                    Application.DoEvents();
                    TOS9200.SendCommand("STOP");
                    TOS9200.Dsr = 0;    //DSRレジスタ初期化
                    TOS9200.Fail = 0;   //failレジスタ初期化
                    TOS9200.SetFunction(spec);//試験設定
                    SetSpecLabelColor(spec);

                    TOS9200.TosStart();


                    spec.計測値 = TOS9200.GetMeasureData() + "Ω"; //計測データの取得

                    switch (spec.ステップ)
                    {
                        case "STEP1":
                            labelBcr絶縁1計測値.Text = spec.計測値;
                            if (TOS9200.ErrorMess != TOS9200.ErrorMessage.正常終了)
                            {
                                labelBcr絶縁1計測値.ForeColor = Color.Red;
                                goto NgStep;
                            }
                            break;
                        case "STEP2":
                            labelBcr絶縁2計測値.Text = spec.計測値;
                            if (TOS9200.ErrorMess != TOS9200.ErrorMessage.正常終了)
                            {
                                labelBcr絶縁2計測値.ForeColor = Color.Red;
                                goto NgStep;
                            }
                            break;
                        case "STEP3":
                            labelBcr絶縁3計測値.Text = spec.計測値;
                            if (TOS9200.ErrorMess != TOS9200.ErrorMessage.正常終了)
                            {
                                labelBcr絶縁3計測値.ForeColor = Color.Red;
                                goto NgStep;
                            }
                            break;

                        default:
                            break;
                    }

                    TOS9200.SendCommand("STOP");

                }
            }
            else if (State.Category == State.CATEGORY.Q890)
            {
                foreach (var spec in State.絶縁抵抗スペックリストQ890)
                {
                    Application.DoEvents();
                    TOS9200.SendCommand("STOP");
                    TOS9200.Dsr = 0;    //DSRレジスタ初期化
                    TOS9200.Fail = 0;   //failレジスタ初期化
                    TOS9200.SetFunction(spec);//試験設定
                    SetSpecLabelColor(spec);

                    TOS9200.TosStart();


                    spec.計測値 = TOS9200.GetMeasureData() + "Ω"; //計測データの取得

                    switch (spec.ステップ)
                    {
                        case "STEP1":
                            labelQ890絶縁1計測値.Text = spec.計測値;
                            if (TOS9200.ErrorMess != TOS9200.ErrorMessage.正常終了)
                            {
                                labelQ890絶縁1計測値.ForeColor = Color.Red;
                                goto NgStep;
                            }
                            break;


                        default:
                            break;
                    }

                    TOS9200.SendCommand("STOP");

                }

            }
            else
            {
                foreach (var spec in State.絶縁抵抗スペックリストAur)
                {
                    Application.DoEvents();
                    TOS9200.SendCommand("STOP");
                    TOS9200.Dsr = 0;    //DSRレジスタ初期化
                    TOS9200.Fail = 0;   //failレジスタ初期化
                    TOS9200.SetFunction(spec);//試験設定
                    SetSpecLabelColor(spec);

                    TOS9200.TosStart();


                    spec.計測値 = TOS9200.GetMeasureData() + "Ω"; //計測データの取得

                    switch (spec.ステップ)
                    {
                        case "STEP1":
                            labelNewAur絶縁1計測値.Text = spec.計測値;
                            if (TOS9200.ErrorMess != TOS9200.ErrorMessage.正常終了)
                            {
                                labelNewAur絶縁1計測値.ForeColor = Color.Red;
                                goto NgStep;
                            }
                            break;
                        case "STEP2":
                            labelNewAur絶縁2計測値.Text = spec.計測値;
                            if (TOS9200.ErrorMess != TOS9200.ErrorMessage.正常終了)
                            {
                                labelNewAur絶縁2計測値.ForeColor = Color.Red;
                                goto NgStep;
                            }
                            break;
                        case "STEP3":
                            labelNewAur絶縁3計測値.Text = spec.計測値;
                            if (TOS9200.ErrorMess != TOS9200.ErrorMessage.正常終了)
                            {
                                labelNewAur絶縁3計測値.ForeColor = Color.Red;
                                goto NgStep;
                            }
                            break;

                        default:
                            break;
                    }

                    TOS9200.SendCommand("STOP");

                }

            }

            SetSpecLabelColor((絶縁抵抗試験スペック)null);//すべてのラベル色を戻すための処理　引数にnullを渡す

            //合格の処理●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●
            Method.InterLock発動();

            TOS9200.SendCommand("STOP");
            State.errorCode = ((int)TOS9200.ErrorMess).ToString("D2");//合格なのでエラーコード00

            //試験データ生成
            var dt = DateTime.Now;
            State.date = dt.ToString("yyyyMMdd,HH:mm:ss");

            State.OperatorName = comboBoxOperatorName.SelectedItem.ToString();
            var logData = State.index + "," + State.model + "," + State.opecode + "," +
                          State.date + "," + State.OperatorName + "," + State.errorCode;


            if (!Method.WriteLogData(logData))
            {
                MessageBox.Show("ログデータ生成に失敗しました\r\n" + "アプリケーションを強制終了します", "警告");
                Environment.Exit(0);
            }


            //PC4内に試験データを保存する
            if (State.Category == State.CATEGORY.BCR)
            {
                var PassData = new List<string>()
                {
                    State.errorCode,
                    State.index,
                    State.model,
                    State.opecode,
                    State.date,
                    State.OperatorName,
                    State.耐電圧スペックリストBcr[0].計測値,
                    State.耐電圧スペックリストBcr[1].計測値,
                    State.絶縁抵抗スペックリストBcr[0].計測値,
                    State.絶縁抵抗スペックリストBcr[1].計測値,
                    State.絶縁抵抗スペックリストBcr[2].計測値,
                    State.耐電圧スペックリストBcr[0].ステップ   + "/一次" + State.耐電圧スペックリストBcr[0].CH1   + "/二次a" + State.耐電圧スペックリストBcr[0].CH2 + "/二次b" + State.耐電圧スペックリストBcr[0].CH3   + "/ダンパ" + State.耐電圧スペックリストBcr[0].CH4   + "/モニタ" + State.耐電圧スペックリストBcr[0].CH5 +"/アース" + State.耐電圧スペックリストBcr[0].CH6   + "/" + State.耐電圧スペックリストBcr[0].印加電圧   + "/" + State.耐電圧スペックリストBcr[0].印加時間   +  "/" + State.耐電圧スペックリストBcr[0].漏れ電流,
                    State.耐電圧スペックリストBcr[1].ステップ   + "/一次" + State.耐電圧スペックリストBcr[1].CH1   + "/二次a" + State.耐電圧スペックリストBcr[1].CH2 + "/二次b" + State.耐電圧スペックリストBcr[1].CH3    + "/ダンパ" + State.耐電圧スペックリストBcr[1].CH4   + "/モニタ" + State.耐電圧スペックリストBcr[1].CH5  +"/アース" + State.耐電圧スペックリストBcr[1].CH6  + "/" + State.耐電圧スペックリストBcr[1].印加電圧   + "/" + State.耐電圧スペックリストBcr[1].印加時間   +  "/" + State.耐電圧スペックリストBcr[1].漏れ電流,
                    State.絶縁抵抗スペックリストBcr[0].ステップ + "/一次" + State.絶縁抵抗スペックリストBcr[0].CH1 + "/二次a" + State.絶縁抵抗スペックリストBcr[0].CH2 + "/二次b" + State.絶縁抵抗スペックリストBcr[0].CH3  + "/ダンパ" + State.絶縁抵抗スペックリストBcr[0].CH4 + "/モニタ" + State.絶縁抵抗スペックリストBcr[0].CH5 +"/アース" + State.絶縁抵抗スペックリストBcr[0].CH6  + "/" + State.絶縁抵抗スペックリストBcr[0].印加電圧 + "/" + State.絶縁抵抗スペックリストBcr[0].印加時間 +  "/" + State.絶縁抵抗スペックリストBcr[0].絶縁抵抗値,
                    State.絶縁抵抗スペックリストBcr[1].ステップ + "/一次" + State.絶縁抵抗スペックリストBcr[1].CH1 + "/二次a" + State.絶縁抵抗スペックリストBcr[1].CH2 + "/二次b" + State.絶縁抵抗スペックリストBcr[1].CH3  + "/ダンパ" + State.絶縁抵抗スペックリストBcr[1].CH4 + "/モニタ" + State.絶縁抵抗スペックリストBcr[1].CH5  +"/アース" + State.絶縁抵抗スペックリストBcr[1].CH6 + "/" + State.絶縁抵抗スペックリストBcr[1].印加電圧 + "/" + State.絶縁抵抗スペックリストBcr[1].印加時間 +  "/" + State.絶縁抵抗スペックリストBcr[1].絶縁抵抗値,
                    State.絶縁抵抗スペックリストBcr[2].ステップ + "/一次" + State.絶縁抵抗スペックリストBcr[2].CH1 + "/二次a" + State.絶縁抵抗スペックリストBcr[2].CH2 + "/二次b" + State.絶縁抵抗スペックリストBcr[2].CH3  + "/ダンパ" + State.絶縁抵抗スペックリストBcr[2].CH4 + "/モニタ" + State.絶縁抵抗スペックリストBcr[2].CH5  +"/アース" + State.絶縁抵抗スペックリストBcr[2].CH6 + "/" + State.絶縁抵抗スペックリストBcr[2].印加電圧 + "/" + State.絶縁抵抗スペックリストBcr[2].印加時間 +  "/" + State.絶縁抵抗スペックリストBcr[2].絶縁抵抗値,
                };

                if (!Method.SaveTestData(PassData))
                {
                    MessageBox.Show("試験データ保存に失敗しました\r\n" + "生産準備 畔上さんを呼んでください！", "警告");
                    Environment.Exit(0);
                }
            }
            else if (State.Category == State.CATEGORY.Q890)
            {
                var PassData = new List<string>()
                {
                    State.errorCode,
                    State.index,
                    State.model,
                    State.opecode,
                    State.date,
                    State.OperatorName,
                    State.耐電圧スペックリストQ890[0].計測値,
                    State.絶縁抵抗スペックリストQ890[0].計測値,
                    "---",
                    "---",
                    State.耐電圧スペックリストQ890[0].ステップ   + "/全端子" + State.耐電圧スペックリストQ890[0].CH1   + "/アース" + State.耐電圧スペックリストQ890[0].CH2   + "/未使用" + State.耐電圧スペックリストQ890[0].CH3   + "/未使用" + State.耐電圧スペックリストQ890[0].CH4   + "/" + State.耐電圧スペックリストQ890[0].印加電圧   + "/" + State.耐電圧スペックリストQ890[0].印加時間   +  "/" + State.耐電圧スペックリストQ890[0].漏れ電流,
                    State.絶縁抵抗スペックリストQ890[0].ステップ + "/全端子" + State.絶縁抵抗スペックリストQ890[0].CH1 + "/アース" + State.絶縁抵抗スペックリストQ890[0].CH2 + "/未使用" + State.絶縁抵抗スペックリストQ890[0].CH3 + "/未使用" + State.絶縁抵抗スペックリストQ890[0].CH4 + "/" + State.絶縁抵抗スペックリストQ890[0].印加電圧 + "/" + State.絶縁抵抗スペックリストQ890[0].印加時間 +  "/" + State.絶縁抵抗スペックリストQ890[0].絶縁抵抗値,
                    "---",
                    "---",
                };

                if (!Method.SaveTestData(PassData))
                {
                    MessageBox.Show("試験データ保存に失敗しました\r\n" + "アプリケーションを強制終了します", "警告");
                    Environment.Exit(0);
                }

            }
            else
            {
                var PassData = new List<string>()
                {
                    State.errorCode,
                    State.index,
                    State.model,
                    State.opecode,
                    State.date,
                    State.OperatorName,
                    State.耐電圧スペックリストAur[0].計測値,
                    State.耐電圧スペックリストAur[1].計測値,
                    State.耐電圧スペックリストAur[2].計測値,
                    State.絶縁抵抗スペックリストAur[0].計測値,
                    State.絶縁抵抗スペックリストAur[1].計測値,
                    State.絶縁抵抗スペックリストAur[2].計測値,
                    State.耐電圧スペックリストAur[0].ステップ   + "/一次" + State.耐電圧スペックリストAur[0].CH1   + "/二次a" + State.耐電圧スペックリストAur[0].CH2 + "/二次b" + State.耐電圧スペックリストAur[0].CH3   + "/ダンパ" + State.耐電圧スペックリストAur[0].CH4   + "/モニタ" + State.耐電圧スペックリストAur[0].CH5 +"/アース" + State.耐電圧スペックリストAur[0].CH6   + "/" + State.耐電圧スペックリストAur[0].印加電圧   + "/" + State.耐電圧スペックリストAur[0].印加時間   +  "/" + State.耐電圧スペックリストAur[0].漏れ電流,
                    State.耐電圧スペックリストAur[1].ステップ   + "/一次" + State.耐電圧スペックリストAur[1].CH1   + "/二次a" + State.耐電圧スペックリストAur[1].CH2 + "/二次b" + State.耐電圧スペックリストAur[1].CH3    + "/ダンパ" + State.耐電圧スペックリストAur[1].CH4   + "/モニタ" + State.耐電圧スペックリストAur[1].CH5  +"/アース" + State.耐電圧スペックリストAur[1].CH6  + "/" + State.耐電圧スペックリストAur[1].印加電圧   + "/" + State.耐電圧スペックリストAur[1].印加時間   +  "/" + State.耐電圧スペックリストAur[1].漏れ電流,
                    State.耐電圧スペックリストAur[2].ステップ   + "/一次" + State.耐電圧スペックリストAur[2].CH1   + "/二次a" + State.耐電圧スペックリストAur[2].CH2 + "/二次b" + State.耐電圧スペックリストAur[2].CH3    + "/ダンパ" + State.耐電圧スペックリストAur[2].CH4   + "/モニタ" + State.耐電圧スペックリストAur[2].CH5  +"/アース" + State.耐電圧スペックリストAur[2].CH6  + "/" + State.耐電圧スペックリストAur[2].印加電圧   + "/" + State.耐電圧スペックリストAur[2].印加時間   +  "/" + State.耐電圧スペックリストAur[2].漏れ電流,
                    State.絶縁抵抗スペックリストAur[0].ステップ + "/一次" + State.絶縁抵抗スペックリストAur[0].CH1 + "/二次a" + State.絶縁抵抗スペックリストAur[0].CH2 + "/二次b" + State.絶縁抵抗スペックリストAur[0].CH3  + "/ダンパ" + State.絶縁抵抗スペックリストAur[0].CH4 + "/モニタ" + State.絶縁抵抗スペックリストAur[0].CH5 +"/アース" + State.絶縁抵抗スペックリストAur[0].CH6  + "/" + State.絶縁抵抗スペックリストAur[0].印加電圧 + "/" + State.絶縁抵抗スペックリストAur[0].印加時間 +  "/" + State.絶縁抵抗スペックリストAur[0].絶縁抵抗値,
                    State.絶縁抵抗スペックリストAur[1].ステップ + "/一次" + State.絶縁抵抗スペックリストAur[1].CH1 + "/二次a" + State.絶縁抵抗スペックリストAur[1].CH2 + "/二次b" + State.絶縁抵抗スペックリストAur[1].CH3  + "/ダンパ" + State.絶縁抵抗スペックリストAur[1].CH4 + "/モニタ" + State.絶縁抵抗スペックリストAur[1].CH5  +"/アース" + State.絶縁抵抗スペックリストAur[1].CH6 + "/" + State.絶縁抵抗スペックリストAur[1].印加電圧 + "/" + State.絶縁抵抗スペックリストAur[1].印加時間 +  "/" + State.絶縁抵抗スペックリストAur[1].絶縁抵抗値,
                    State.絶縁抵抗スペックリストAur[2].ステップ + "/一次" + State.絶縁抵抗スペックリストAur[2].CH1 + "/二次a" + State.絶縁抵抗スペックリストAur[2].CH2 + "/二次b" + State.絶縁抵抗スペックリストAur[2].CH3  + "/ダンパ" + State.絶縁抵抗スペックリストAur[2].CH4 + "/モニタ" + State.絶縁抵抗スペックリストAur[2].CH5  +"/アース" + State.絶縁抵抗スペックリストAur[2].CH6 + "/" + State.絶縁抵抗スペックリストAur[2].印加電圧 + "/" + State.絶縁抵抗スペックリストAur[2].印加時間 +  "/" + State.絶縁抵抗スペックリストAur[2].絶縁抵抗値,
                };

                if (!Method.SaveTestData(PassData))
                {
                    MessageBox.Show("試験データ保存に失敗しました\r\n" + "生産準備 畔上さんを呼んでください！", "警告");
                    Environment.Exit(0);
                }
            }

            General.PlaySound2(Constants.SoundPass);

            labelDecision.Text = "PASS";
            labelDecision.ForeColor = Color.LightSeaGreen;
            labelDanger.BackColor = Color.MistyRose;

            pictureBoxWarning.Visible = false;
            labelMessage.Location = new Point(10, 15);
            labelMessage.Size = new Size(965, 75);
            labelMessage.Text = Constants.MessRemove;
            labelMessage.BackColor = SystemColors.GradientActiveCaption;

            timerLbMessage.Start();

            Flags.FlagTest = false;
            labelMessage.Text = "製品を取り外してください";
            while (true)
            {
                Application.DoEvents();
                Method.io.ReadInputData(EPX64R.PortName.P7);
                if ((byte)(Method.io.P7InputData & 0x13) == 0x13)
                    break;
            }
            ClearForm();
            return;

        NgStep:

            TOS9200.SendCommand("STOP");
            State.errorCode = ((int)TOS9200.ErrorMess).ToString("D2");//不合格なのでエラーコード00以外

            Method.InterLock発動();


            //試験データ生成
            dt = DateTime.Now;
            State.date = dt.ToString("yyyyMMdd,HH:mm:ss");

            State.OperatorName = comboBoxOperatorName.SelectedItem.ToString();
            logData = State.index + "," + State.model + "," + State.opecode + "," +
                          State.date + "," + State.OperatorName + "," + State.errorCode;


            if (!Method.WriteLogData(logData))
            {
                MessageBox.Show("ログデータ生成に失敗しました\r\n" + "アプリケーションを強制終了します", "警告");
                Environment.Exit(0);
            }


            if (State.Category == State.CATEGORY.BCR)
            {
                var FailData = new List<string>()
                {
                    State.errorCode,
                    State.index,
                    State.model,
                    State.opecode,
                    State.date,
                    State.OperatorName,
                    State.耐電圧スペックリストBcr[0].計測値,
                    State.耐電圧スペックリストBcr[1].計測値,
                    State.絶縁抵抗スペックリストBcr[0].計測値,
                    State.絶縁抵抗スペックリストBcr[1].計測値,
                    State.絶縁抵抗スペックリストBcr[2].計測値,
                    State.耐電圧スペックリストBcr[0].ステップ   + "/一次" + State.耐電圧スペックリストBcr[0].CH1   + "/二次a" + State.耐電圧スペックリストBcr[0].CH2   + "/二次b" + State.耐電圧スペックリストBcr[0].CH3  + "/ダンパ" + State.耐電圧スペックリストBcr[0].CH4   + "/モニタ" + State.耐電圧スペックリストBcr[0].CH5  + "/アース" + State.耐電圧スペックリストBcr[0].CH6 + "/" + State.耐電圧スペックリストBcr[0].印加電圧   + "/" + State.耐電圧スペックリストBcr[0].印加時間   +  "/" + State.耐電圧スペックリストBcr[0].漏れ電流,
                    State.耐電圧スペックリストBcr[1].ステップ   + "/一次" + State.耐電圧スペックリストBcr[1].CH1   + "/二次a" + State.耐電圧スペックリストBcr[1].CH2   + "/二次b" + State.耐電圧スペックリストBcr[1].CH3   + "/ダンパ" + State.耐電圧スペックリストBcr[1].CH4   + "/モニタ" + State.耐電圧スペックリストBcr[1].CH5  + "/アース" + State.耐電圧スペックリストBcr[1].CH6  + "/" + State.耐電圧スペックリストBcr[1].印加電圧   + "/" + State.耐電圧スペックリストBcr[1].印加時間   +  "/" + State.耐電圧スペックリストBcr[1].漏れ電流,
                    State.絶縁抵抗スペックリストBcr[0].ステップ + "/一次" + State.絶縁抵抗スペックリストBcr[0].CH1 + "/二次a" + State.絶縁抵抗スペックリストBcr[0].CH2 + "/二次b" + State.絶縁抵抗スペックリストBcr[0].CH3   + "/ダンパ" + State.絶縁抵抗スペックリストBcr[0].CH4 + "/モニタ" + State.絶縁抵抗スペックリストBcr[0].CH5 + "/アース" + State.絶縁抵抗スペックリストBcr[0].CH6 + "/" + State.絶縁抵抗スペックリストBcr[0].印加電圧 + "/" + State.絶縁抵抗スペックリストBcr[0].印加時間 +  "/" + State.絶縁抵抗スペックリストBcr[0].絶縁抵抗値,
                    State.絶縁抵抗スペックリストBcr[1].ステップ + "/一次" + State.絶縁抵抗スペックリストBcr[1].CH1 + "/二次a" + State.絶縁抵抗スペックリストBcr[1].CH2 + "/二次b" + State.絶縁抵抗スペックリストBcr[1].CH3  + "/ダンパ" + State.絶縁抵抗スペックリストBcr[1].CH4 + "/モニタ" + State.絶縁抵抗スペックリストBcr[1].CH5 + "/アース" + State.絶縁抵抗スペックリストBcr[1].CH6 + "/" + State.絶縁抵抗スペックリストBcr[1].印加電圧 + "/" + State.絶縁抵抗スペックリストBcr[1].印加時間 +  "/" + State.絶縁抵抗スペックリストBcr[1].絶縁抵抗値,
                    State.絶縁抵抗スペックリストBcr[2].ステップ + "/一次" + State.絶縁抵抗スペックリストBcr[2].CH1 + "/二次a" + State.絶縁抵抗スペックリストBcr[2].CH2 + "/二次b" + State.絶縁抵抗スペックリストBcr[2].CH3  + "/ダンパ" + State.絶縁抵抗スペックリストBcr[2].CH4 + "/モニタ" + State.絶縁抵抗スペックリストBcr[2].CH5 + "/アース" + State.絶縁抵抗スペックリストBcr[2].CH6 + "/" + State.絶縁抵抗スペックリストBcr[2].印加電圧 + "/" + State.絶縁抵抗スペックリストBcr[2].印加時間 +  "/" + State.絶縁抵抗スペックリストBcr[2].絶縁抵抗値,

                };

                if (!Method.SaveTestData(FailData))
                {
                    MessageBox.Show("試験データ保存に失敗しました\r\n" + "アプリケーションを強制終了します", "警告");
                    Environment.Exit(0);
                }
            }
            else if (State.Category == State.CATEGORY.Q890)
            {
                var FailData = new List<string>()
                {
                    State.errorCode,
                    State.index,
                    State.model,
                    State.opecode,
                    State.date,
                    State.OperatorName,
                    State.耐電圧スペックリストQ890[0].計測値,
                    State.絶縁抵抗スペックリストQ890[0].計測値,
                    "---",
                    "---",
                    State.耐電圧スペックリストQ890[0].ステップ   + "/一次" + State.耐電圧スペックリストQ890[0].CH1   + "/二次" + State.耐電圧スペックリストQ890[0].CH2   + "/ダンパ" + State.耐電圧スペックリストQ890[0].CH3   + "/モニタ" + State.耐電圧スペックリストQ890[0].CH4   + "/" + State.耐電圧スペックリストQ890[0].印加電圧   + "/" + State.耐電圧スペックリストQ890[0].印加時間   +  "/" + State.耐電圧スペックリストQ890[0].漏れ電流,
                    State.絶縁抵抗スペックリストQ890[0].ステップ + "/一次" + State.絶縁抵抗スペックリストQ890[0].CH1 + "/二次" + State.絶縁抵抗スペックリストQ890[0].CH2 + "/ダンパ" + State.絶縁抵抗スペックリストQ890[0].CH3 + "/モニタ" + State.絶縁抵抗スペックリストQ890[0].CH4 + "/" + State.絶縁抵抗スペックリストQ890[0].印加電圧 + "/" + State.絶縁抵抗スペックリストQ890[0].印加時間 +  "/" + State.絶縁抵抗スペックリストQ890[0].絶縁抵抗値,
                    "---",
                    "---",

                };

                if (!Method.SaveTestData(FailData))
                {
                    MessageBox.Show("試験データ保存に失敗しました\r\n" + "アプリケーションを強制終了します", "警告");
                    Environment.Exit(0);
                }
            }
            else
            {
                var FailData = new List<string>()
                {
                    State.errorCode,
                    State.index,
                    State.model,
                    State.opecode,
                    State.date,
                    State.OperatorName,
                    State.耐電圧スペックリストAur[0].計測値,
                    State.耐電圧スペックリストAur[1].計測値,
                    State.耐電圧スペックリストAur[2].計測値,
                    State.絶縁抵抗スペックリストAur[0].計測値,
                    State.絶縁抵抗スペックリストAur[1].計測値,
                    State.絶縁抵抗スペックリストAur[2].計測値,
                    State.耐電圧スペックリストAur[0].ステップ   + "/一次" + State.耐電圧スペックリストAur[0].CH1   + "/二次a" + State.耐電圧スペックリストAur[0].CH2   + "/二次b" + State.耐電圧スペックリストAur[0].CH3  + "/ダンパ" + State.耐電圧スペックリストAur[0].CH4   + "/モニタ" + State.耐電圧スペックリストAur[0].CH5  + "/アース" + State.耐電圧スペックリストAur[0].CH6 + "/" + State.耐電圧スペックリストAur[0].印加電圧   + "/" + State.耐電圧スペックリストAur[0].印加時間   +  "/" + State.耐電圧スペックリストAur[0].漏れ電流,
                    State.耐電圧スペックリストAur[1].ステップ   + "/一次" + State.耐電圧スペックリストAur[1].CH1   + "/二次a" + State.耐電圧スペックリストAur[1].CH2   + "/二次b" + State.耐電圧スペックリストAur[1].CH3   + "/ダンパ" + State.耐電圧スペックリストAur[1].CH4   + "/モニタ" + State.耐電圧スペックリストAur[1].CH5  + "/アース" + State.耐電圧スペックリストAur[1].CH6  + "/" + State.耐電圧スペックリストAur[1].印加電圧   + "/" + State.耐電圧スペックリストAur[1].印加時間   +  "/" + State.耐電圧スペックリストAur[1].漏れ電流,
                    State.耐電圧スペックリストAur[2].ステップ   + "/一次" + State.耐電圧スペックリストAur[2].CH1   + "/二次a" + State.耐電圧スペックリストAur[2].CH2   + "/二次b" + State.耐電圧スペックリストAur[2].CH3   + "/ダンパ" + State.耐電圧スペックリストAur[2].CH4   + "/モニタ" + State.耐電圧スペックリストAur[2].CH5  + "/アース" + State.耐電圧スペックリストAur[2].CH6  + "/" + State.耐電圧スペックリストAur[2].印加電圧   + "/" + State.耐電圧スペックリストAur[2].印加時間   +  "/" + State.耐電圧スペックリストAur[2].漏れ電流,
                    State.絶縁抵抗スペックリストAur[0].ステップ + "/一次" + State.絶縁抵抗スペックリストAur[0].CH1 + "/二次a" + State.絶縁抵抗スペックリストAur[0].CH2 + "/二次b" + State.絶縁抵抗スペックリストAur[0].CH3   + "/ダンパ" + State.絶縁抵抗スペックリストAur[0].CH4 + "/モニタ" + State.絶縁抵抗スペックリストAur[0].CH5 + "/アース" + State.絶縁抵抗スペックリストAur[0].CH6 + "/" + State.絶縁抵抗スペックリストAur[0].印加電圧 + "/" + State.絶縁抵抗スペックリストAur[0].印加時間 +  "/" + State.絶縁抵抗スペックリストAur[0].絶縁抵抗値,
                    State.絶縁抵抗スペックリストAur[1].ステップ + "/一次" + State.絶縁抵抗スペックリストAur[1].CH1 + "/二次a" + State.絶縁抵抗スペックリストAur[1].CH2 + "/二次b" + State.絶縁抵抗スペックリストAur[1].CH3  + "/ダンパ" + State.絶縁抵抗スペックリストAur[1].CH4 + "/モニタ" + State.絶縁抵抗スペックリストAur[1].CH5 + "/アース" + State.絶縁抵抗スペックリストAur[1].CH6 + "/" + State.絶縁抵抗スペックリストAur[1].印加電圧 + "/" + State.絶縁抵抗スペックリストAur[1].印加時間 +  "/" + State.絶縁抵抗スペックリストAur[1].絶縁抵抗値,
                    State.絶縁抵抗スペックリストAur[2].ステップ + "/一次" + State.絶縁抵抗スペックリストAur[2].CH1 + "/二次a" + State.絶縁抵抗スペックリストAur[2].CH2 + "/二次b" + State.絶縁抵抗スペックリストAur[2].CH3  + "/ダンパ" + State.絶縁抵抗スペックリストAur[2].CH4 + "/モニタ" + State.絶縁抵抗スペックリストAur[2].CH5 + "/アース" + State.絶縁抵抗スペックリストAur[2].CH6 + "/" + State.絶縁抵抗スペックリストAur[2].印加電圧 + "/" + State.絶縁抵抗スペックリストAur[2].印加時間 +  "/" + State.絶縁抵抗スペックリストAur[2].絶縁抵抗値,

                };

                if (!Method.SaveTestData(FailData))
                {
                    MessageBox.Show("試験データ保存に失敗しました\r\n" + "アプリケーションを強制終了します", "警告");
                    Environment.Exit(0);
                }
            }






            General.PlaySound2(Constants.SoundFail);

            labelDecision.Text = "FAIL";
            labelDecision.ForeColor = Color.Red;
            labelDanger.BackColor = Color.MistyRose;
            labelErrorMessage.Text = TOS9200.ErrorMess.ToString();

            Flags.FlagTest = false;

            pictureBoxWarning.Visible = false;
            labelMessage.Location = new Point(10, 15);
            labelMessage.Size = new Size(965, 75);
            labelMessage.Text = Constants.MessRemove;
            labelMessage.BackColor = SystemColors.GradientActiveCaption;

            labelMessage.Text = "製品を取り外してください";
            timerLbMessage.Start();
            while (true)
            {
                Application.DoEvents();
                Method.io.ReadInputData(EPX64R.PortName.P7);
                if ((byte)(Method.io.P7InputData & 0x13) == 0x13)
                    break;
            }

            SetSpecLabelColor((耐電圧試験スペック)null);//すべてのラベル色を戻すための処理
            SetSpecLabelColor((絶縁抵抗試験スペック)null);//すべてのラベル色を戻すための処理　引数にnullを渡す

            ClearForm();
            return;

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        bool LedSw = false;
        private void TimerLed_Tick(object sender, EventArgs e)
        {
            LedSw = !LedSw;
            Method.SetLed(LedSw);
        }
    }
}
