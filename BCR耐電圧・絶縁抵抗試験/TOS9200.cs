using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Ports;
using System.Windows.Forms;
using System.Timers;
using System.Threading;

namespace BCR耐電圧_絶縁抵抗試験
{
    public static class TOS9200
    {
        
        //デバイスステータスレジスタの定義
        public const int DSR_READY = 1;
        public const int DSR_INVSET = 2;
        public const int DSR_TEST = 4;
        public const int DSR_TESTON = 8;
        public const int DSR_PASS = 16;
        public const int DSR_FAIL = 32;
        public const int DSR_STOP = 64;
        public const int DSR_PROTECTION = 128;
                
        //フェイルレジスタの定義
        public const int CONTACT_FAIL = 1;
        public const int LOWER_FAIL = 2;
        public const int UPPER_FAIL = 4;

        //プロテクションレジスタ１の定義
        public const int OHP = 1;
        public const int OHPT = 2;
        public const int LVP = 16;
        public const int OVLD = 32;
        
        //プロテクションレジスタ２の定義
        public const int LCK = 4; //インターロック
        public const int REN = 32;

        public enum ErrorMessage
        { 
            正常終了,
            強制停止,            
            インターロック作動,
            コンタクトエラー,
            絶縁抵抗値下限異常,
            漏洩電流上限異常,
            その他異常,
            TOS9200設定異常,
        }


        //静的変数の宣言
        public static SerialPort Port;

        public static int Dsr; //デバイスステータスレジスタの値
        public static int Fail; //フェイルレジスタの値
        public static string Result; //試験結果
        public static ErrorMessage ErrorMess;

        //プロパティ
        public static string RecieveData { get; private set; }  //TOS9200から受信した生データ


        //**************************************************************************
        //TOS9200のイニシャライズ
        //**************************************************************************
        public static bool InitTOS9200(PORT_NAME pName)
        {                
            try
            {
                //TOS9200用のシリアルポート設定
                Port = new SerialPort();
                Port.PortName = pName.ToString();
                Port.BaudRate = 9600;
                Port.DataBits = 8;
                Port.Parity = System.IO.Ports.Parity.None;
                Port.StopBits = System.IO.Ports.StopBits.One;
                Port.DtrEnable = true;
                Port.NewLine = ("\r\n");

                //TOS9200用のポートが開いているかどうかの判定
                if (Port.IsOpen == false) Port.Open();
                return true;
            }
            catch
            {
                return false;
            }

        }


        //**************************************************************************
        //TOS9200からの受信データを読み取る
        //引数：指定時間（ｍｓｅｃ）
        //戻値：bool値（正常：true　異常：false）
        //**************************************************************************
        public static bool ReadRecieveData(int time)
        {
            RecieveData = "";   //初期化
            Port.ReadTimeout = time;
            
            try
            {
                RecieveData = Port.ReadTo("\r\n");
                return true;
            }
            catch
            {
                return false;
            }
        }


        //**************************************************************************
        //デバイスへのプログラムメッセージ送信
        //引数：コマンドメッセージ
        //戻値：bool
        //**************************************************************************
        public static bool SendCommand(string data, bool check = true)
        {
            try
            {
                Port.DiscardInBuffer();//データ送信前に受信バッファのクリア
                Port.WriteLine(data);

                if (check)
                {
                    if (!ReadRecieveData(1000)) return false;
                    if (RecieveData.IndexOf("OK") != 0) return false;
                    return true;
                }

                General.Wait(500);
                return true;
            }
            catch
            {
                return false;
            }
        }

        //**************************************************************************
        //デバイスへのプログラムメッセージ送信
        //引数：クエリメッセージ
        //戻値：bool
        //**************************************************************************
        public static bool SendQuery(string data)
        {
            try
            {
                Port.DiscardInBuffer();//データ送信前に受信バッファのクリア
                Port.WriteLine(data);
                return true;
            }
            catch
            {
                return false;
            }
        }

        //**************************************************************************
        //TOS9200の試験設定
        //引数：なし
        //戻値：なし
        //**************************************************************************

        public static bool SetFunction(耐電圧試験スペック spec, int vol = 1, string コンタクトチェック = "OFF")//vol（ブザー音量）は1(Min)～3(Max)で設定すること
        {
            //設定
            SendCommand("FUNCTION 0");  //ACW画面に設定
            if (!Query("FUN?", (data) => data == "0")) goto 設定異常;

            SendCommand("A:FREQ 50");  //周波数50Hzに設定(固定)
            if (!Query("A:FREQ?", (data) => data == "50")) goto 設定異常;

            SendCommand("A:TES " + spec.印加電圧.ToString("0.00E0"));//試験電圧の設定
            if ( !Query("A:TES?",  (data) => Double.Parse(data) == spec.印加電圧) ) goto 設定異常;
            
            SendCommand("A:UPP " + (spec.漏れ電流 * 0.001).ToString("0.00E0"));  //上限基準値の設定
            if ( !Query("A:UPP?", (data) => Double.Parse(data) == (spec.漏れ電流 * 0.001) ) ) goto 設定異常;
            
            SendCommand("A:TIM " + spec.印加時間.ToString("F1") + ",1");  //試験時間の設定
            if ( !Query("A:TIM?", (data) => data == spec.印加時間.ToString("F1") + ",1" ) ) goto 設定異常;
            
            SendCommand("PASSHOLD HOLD");  //試験時間の設定

            SendCommand("A:SCAN 1," + spec.CH1);
            if (!Query("A:SCAN?1", (data) => data == spec.CH1)) goto 設定異常;
            
            SendCommand("A:SCAN 2," + spec.CH2);
            if (!Query("A:SCAN?2", (data) => data == spec.CH2)) goto 設定異常;
            
            SendCommand("A:SCAN 3," + spec.CH3);
            if (!Query("A:SCAN?3", (data) => data == spec.CH3)) goto 設定異常;

            SendCommand("A:SCAN 4," + spec.CH4);
            if (!Query("A:SCAN?4", (data) => data == spec.CH4)) goto 設定異常;

            SendCommand("A:CCH " + コンタクトチェック);

            SendCommand("DSE #HFF");//デバイスステータスイネーブルレジスタをFFに設定
            if (!Query("DSE?", (data) => data == "255")) goto 設定異常;

            SendCommand("BVOL " + vol.ToString());

            return true;

        設定異常:
            ErrorMess = ErrorMessage.TOS9200設定異常;
        return false;                

        }

        public static bool SetFunction(絶縁抵抗試験スペック spec, int vol = 1, string コンタクトチェック = "OFF")//vol（ブザー音量）は1(Min)～3(Max)で設定すること
        {

            SendCommand("FUNCTION 2");  //IR画面に設定
            if (!Query("FUN?", (data) => data == "2")) goto 設定異常;

            SendCommand("I:TES " + spec.印加電圧.ToString("F0"));  //試験電圧の設定
            if (!Query("I:TES?", (data) => Double.Parse(data) == spec.印加電圧 ) ) goto 設定異常;

            SendCommand("I:LOW " + (spec.絶縁抵抗値 * 1000000).ToString() + ",1");  //下限基準値の設定
            if (!Query("I:LOW?", (data) => {
                int i = data.IndexOf(",1");
                if(i < 0) return false;
                return Double.Parse(data.Substring(0, i)) == spec.絶縁抵抗値 * 1000000;
                })
            ) goto 設定異常;

            SendCommand("I:TIM " + spec.印加時間.ToString("F1") + ",1");  //試験時間の設定
            if (!Query("I:TIM?", (data) => data == spec.印加時間.ToString("F1") + ",1")) goto 設定異常;
            
            SendCommand("PASSHOLD HOLD");  //試験時間の設定

            SendCommand("I:SCAN 1," + spec.CH1);
            if (!Query("I:SCAN?1", (data) => data == spec.CH1)) goto 設定異常;
            
            SendCommand("I:SCAN 2," + spec.CH2);
            if (!Query("I:SCAN?2", (data) => data == spec.CH2)) goto 設定異常;
            
            SendCommand("I:SCAN 3," + spec.CH3);
            if (!Query("I:SCAN?3", (data) => data == spec.CH3)) goto 設定異常;

            SendCommand("I:SCAN 4," + spec.CH4);
            if (!Query("I:SCAN?4", (data) => data == spec.CH4)) goto 設定異常;

            SendCommand("I:CCH " + コンタクトチェック);

            SendCommand("DSE #HFF");//デバイスステータスイネーブルレジスタをFFに設定（）プロテクション中は無効！
            if (!Query("DSE?", (data) => data == "255")) goto 設定異常;
   
            SendCommand("BVOL " + vol.ToString());

            return true;

        設定異常:
            ErrorMess = ErrorMessage.TOS9200設定異常;
            return false;   
        }

        //**************************************************************************
        //ＣＯＭポートを閉じる処理
        //引数：なし
        //戻値：bool
        //**************************************************************************
        public static bool CloseComPort()
        {
            try
            {
                //TOS9200用のポートが開いているかどうかの判定
                if (Port.IsOpen)　Port.Close();
                return true;
            }
            catch
            {
                return false;
            }
        }
        //**************************************************************************
        //ファンクション問い合わせ
        //引数：なし
        //戻値：bool
        //**************************************************************************
        public static bool Query(string クエリ, Func<string, bool> 判定)
        {
            if(!SendQuery(クエリ)) return false;
            ReadRecieveData(1000);
            return 判定(RecieveData);
        }

        //**************************************************************************
        //TOS9200 DSRレジスタのチェック
        //引数：なし
        //戻値：bool
        //**************************************************************************
        public static void CheckDsrRegister()
        {
            switch (Dsr)
            {
                case DSR_PASS:
                    ErrorMess = ErrorMessage.正常終了; break;

                case DSR_STOP:
                case DSR_READY:
                    ErrorMess = ErrorMessage.強制停止; break;
                    
                case DSR_PROTECTION:
                     ErrorMess = ErrorMessage.インターロック作動; break;
                    
                case DSR_FAIL:
                    SendQuery("FAIL?");
                    ReadRecieveData(500);
                    Fail = Int32.Parse(RecieveData);
                    switch(Fail)
                    {
                        case CONTACT_FAIL:
                            ErrorMess = ErrorMessage.コンタクトエラー; break;
                        case LOWER_FAIL:
                            ErrorMess = ErrorMessage.絶縁抵抗値下限異常; break;
                        case UPPER_FAIL:
                            ErrorMess = ErrorMessage.漏洩電流上限異常; break;
                    }
                    break;
                
                default:
                    ErrorMess = ErrorMessage.その他異常; break;
            }

        }

        //**************************************************************************
        //TOS9200 試験開始処理
        //引数：なし
        //戻値：bool値（正常終了：true　インターロック異常：false）
        //**************************************************************************
        public static void TosStart()
        {
            //インターロックが解除されているかどうかの判定
            SendCommand("STOP"); //インターロック解除
            SendQuery("PROT?");
            ReadRecieveData(1000);
            if (RecieveData != "0,0")
            {
                SendCommand("STOP");//念のため停止
                ErrorMess = ErrorMessage.インターロック作動;
                return;
            }

            //レディ状態になったので試験スタート
            SendCommand("START");  //試験開始

            for (; ; ) //試験が終わるまで、DSRレジスタの値と測定値を読み出す
            {
                Application.DoEvents();
                SendQuery("DSR?");
                ReadRecieveData(500);
                Dsr = Int32.Parse(RecieveData);

                SendQuery("MON?");
                ReadRecieveData(500);
                Result = RecieveData;

                if (Dsr == DSR_FAIL || Dsr == DSR_PASS || Dsr == DSR_READY || Dsr == DSR_STOP || Dsr == DSR_PROTECTION) break;
            }

            CheckDsrRegister();

        }


        //**************************************************************************
        //テスト結果から、モニタ電流Max or モニタ抵抗Minを取り出す
        //引数：なし
        //戻値：なし
        //**************************************************************************
        public static string GetMeasureData()
        {
            if (ErrorMess != ErrorMessage.正常終了 && ErrorMess != ErrorMessage.漏洩電流上限異常
                && ErrorMess != ErrorMessage.絶縁抵抗値下限異常) return "--";
            
            try
            {
                int i = 0;

                i = Result.IndexOf(",");//1ヶ目の","の位置
                i = Result.IndexOf(",", i + 1);//2ヶ目の","の位置
                i = Result.IndexOf(",", i + 1);//3ヶ目の","の位置

                int start = i + 1;  //3ヶ目の","の次の文字をスタートとする

                i = Result.IndexOf(",", i + 1);//4ヶ目の","の位置

                int end = i - 1;  //4ヶ目の","の前の文字をエンドとする

                string buff1 = Result.Substring(start, end - start + 1);

                string 仮数 = buff1.Substring(0, buff1.IndexOf("E"));     //仮数部取り出し
                string 指数 = buff1.Substring(buff1.IndexOf("E"));          //指数部取り出し

                string 単位 = "";

                switch (指数)
                {
                    case "E-3":
                        単位 = "m";
                        break;

                    case "E-6":
                        単位 = "μ";
                        break;

                    case "E-9":
                        単位 = "p";
                        break;

                    case "E3":
                        単位 = "K";
                        break;

                    case "E6":
                        単位 = "M";
                        break;

                    case "E9":
                        単位 = "G";
                        break;
                    default:
                        break;
                }


                return 仮数 + 単位;
            }
            catch
            {
                return "--";

            }
            
        
        }
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    }
}
