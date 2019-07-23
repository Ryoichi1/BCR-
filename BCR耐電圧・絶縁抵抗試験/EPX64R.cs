using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

/// <summary>
/// EPX-64R API library definitions
/// </summary>

namespace BCR耐電圧_絶縁抵抗試験
{

    public class EPX64R
    {
        // Function definitions
        [DllImport("EPX64R.dll")]
        public static extern int EPX64R_GetNumberOfDevices(ref int Number);
        [DllImport("EPX64R.dll")]
        public static extern int EPX64R_GetSerialNumber(int Index, ref int SerialNumber);
        [DllImport("EPX64R.dll")]
        public static extern int EPX64R_Open(ref System.IntPtr Handle);
        [DllImport("EPX64R.dll")]
        public static extern int EPX64R_OpenBySerialNumber(int SerialNumber, ref System.IntPtr Handle);
        [DllImport("EPX64R.dll")]
        public static extern int EPX64R_Close(System.IntPtr Handle);
        [DllImport("EPX64R.dll")]
        public static extern int EPX64R_SetDirection(System.IntPtr Handle, byte Direction);
        [DllImport("EPX64R.dll")]
        public static extern int EPX64R_GetDirection(System.IntPtr Handle, ref byte Direction);
        [DllImport("EPX64R.dll")]
        public static extern int EPX64R_OutputPort(System.IntPtr Handle, byte Port, byte Value);
        [DllImport("EPX64R.dll")]
        public static extern int EPX64R_InputPort(System.IntPtr Handle, byte Port, ref byte Value);
        [DllImport("EPX64R.dll")]
        public static extern int EPX64R_SetPortOutputBuffer(System.IntPtr Handle, byte Port, byte Value);
        [DllImport("EPX64R.dll")]
        public static extern int EPX64R_GetPortInputBuffer(System.IntPtr Handle, byte Port, ref byte Value);
        [DllImport("EPX64R.dll")]
        public static extern int EPX64R_SetInputLatch(System.IntPtr Handle);
        [DllImport("EPX64R.dll")]
        public static extern int EPX64R_SetOutputLatch(System.IntPtr Handle);
        
        //クラス変数（定数）宣言 
        // Device status (Return codes) ※const メンバーはクラスに属します（静的変数と同じ扱い）
        public const int EPX64R_OK = 0;
        public const int EPX64R_INVALID_HANDLE = 1;
        public const int EPX64R_DEVICE_NOT_FOUND = 2;
        public const int EPX64R_DEVICE_NOT_OPENED = 3;
        public const int EPX64R_OTHER_ERROR = 4;
        public const int EPX64R_COMMUNICATION_ERROR = 5;
        public const int EPX64R_INVALID_PARAMETER = 6;

        // Constants　※const メンバーはクラスに属します（静的変数と同じ扱い）
        public const byte EPX64R_PORT0 = 0x01;
        public const byte EPX64R_PORT1 = 0x02;
        public const byte EPX64R_PORT2 = 0x04;
        public const byte EPX64R_PORT3 = 0x08;
        public const byte EPX64R_PORT4 = 0x10;
        public const byte EPX64R_PORT5 = 0x20;
        public const byte EPX64R_PORT6 = 0x40;
        public const byte EPX64R_PORT7 = 0x80;


        //列挙型の宣言
        public enum PortName
        { 
            P0, P1, P2, P3, P4, P5, P6, P7,
        }

        public enum BitName
        {
            b0, b1, b2, b3, b4, b5, b6, b7,
        }

        public enum OutData
        {
            H, L,
        }

        //インスタンス変数宣言（生成するインスタンス固有の値）
        public System.IntPtr hDevice;   //Device Handle（ポインタへのポインタ）
        public int Status;              // Device Status (Return Code)
        public int Number;              // Devices Number
        public byte Direction;          //各ポート 入力or出力の設定用パラメータ
        public byte Port;               //入出力のポート指定用パラメータ
        public byte InputValue;         //指定ポートから読み込んだデータ

        public byte P0InputData; //ポート0の入力データバッファ
        public byte P1InputData; //ポート1の入力データバッファ
        public byte P2InputData; //ポート2の入力データバッファ
        public byte P3InputData; //ポート3の入力データバッファ
        public byte P4InputData; //ポート4の入力データバッファ
        public byte P5InputData; //ポート5の入力データバッファ
        public byte P6InputData; //ポート6の入力データバッファ
        public byte P7InputData; //ポート7の入力データバッファ

        public byte p0Outdata; //現在のポート０に出力されているデータ（電源入ってるかチェックするためpublicにしておく）
        public byte p1Outdata; //現在のポート１に出力されているデータ
        public byte p2Outdata; //現在のポート２に出力されているデータ
        public byte p3Outdata; //現在のポート３に出力されているデータ
        public byte p4Outdata; //現在のポート４に出力されているデータ
        public byte p5Outdata; //現在のポート５に出力されているデータ
        public byte p6Outdata; //現在のポート６に出力されているデータ
        public byte p7Outdata; //現在のポート７に出力されているデータ

        //インスタンスコンストラクタ
        public EPX64R()
        {
            //インスタンス変数の初期化
            hDevice = System.IntPtr.Zero;   // Device Handle
            Status = 0; // Device Status (Return Code)
            Number = 0; // Devices Number
            Direction = 0;
            Port = 0;
            InputValue = 0;

            P0InputData = 0; //現在のポート０に出力されているデータ
            P1InputData = 0; //現在のポート１に出力されているデータ
            P2InputData = 0; //現在のポート２に出力されているデータ
            P3InputData = 0; //現在のポート３に出力されているデータ
            P4InputData = 0; //現在のポート４に出力されているデータ
            P5InputData = 0; //現在のポート５に出力されているデータ
            P6InputData = 0; //現在のポート６に出力されているデータ
            P7InputData = 0; //現在のポート７に出力されているデータ

            p0Outdata = 0; //現在のポート０に出力されているデータ
            p1Outdata = 0; //現在のポート１に出力されているデータ
            p2Outdata = 0; //現在のポート２に出力されているデータ
            p3Outdata = 0; //現在のポート３に出力されているデータ
            p4Outdata = 0; //現在のポート４に出力されているデータ
            p5Outdata = 0; //現在のポート５に出力されているデータ
            p6Outdata = 0; //現在のポート６に出力されているデータ
            p7Outdata = 0; //現在のポート７に出力されているデータ
        }


        //**************************************************************************
        //EPX64の初期化
        //引数：なし
        //戻値：なし ※異常時は強制終了
        //**************************************************************************
        public void InitEpx64r(byte directionData)
        {
            try
            {

                // Device Number
                this.Status = EPX64R.EPX64R_GetNumberOfDevices(ref this.Number);
                if (this.Status != EPX64R.EPX64R_OK)
                {
                    MessageBox.Show("EPX64R_GetNumberOfDevices() Error" + "\n" + "強制終了します", "IOﾎﾞｰﾄﾞ異常");
                    Environment.Exit(0);
                }
                if (this.Number == 0)
                {
                    MessageBox.Show("DeviceNumber=0" + "\n" + "強制終了します", "IOﾎﾞｰﾄﾞ異常");
                    Environment.Exit(0);
                }

                // Device Open
                this.Status = EPX64R.EPX64R_Open(ref this.hDevice);
                if (this.Status != EPX64R.EPX64R_OK)
                {
                    MessageBox.Show("EPX64R_Open() Error" + "\n" + "強制終了します", "IOﾎﾞｰﾄﾞ異常");
                    Environment.Exit(0);
                }

                this.Direction = directionData; // Direction 0=入力 1=出力
                this.Status = EPX64R.EPX64R_SetDirection(this.hDevice, this.Direction);
                if (this.Status != EPX64R.EPX64R_OK)
                {
                    EPX64R.EPX64R_Close(this.hDevice);   // Device Close
                    MessageBox.Show("EPX64R_SetDirection() Error" + "\n" + "強制終了します", "IOﾎﾞｰﾄﾞ異常");
                    return;
                }
            }
            catch
            {
                //EPX64R.EPX64R_Close(this.hDevice);   // Device Close
                MessageBox.Show("EPX64R_SetDirection() Error" + "\n" + "強制終了します", "IOﾎﾞｰﾄﾞ異常");
                Environment.Exit(0);
                return;
            }
        }

        //****************************************************************************
        //メソッド名：ReadInputData（指定ポートのデータを取り込む）
        //引数：ポート名（"P0"、"P1"・・・"P7"）
        //戻り値：取り込んだデータ正常０、異常１
        //****************************************************************************
        public int ReadInputData(PortName pName)
        {
            switch (pName)
            {
                case PortName.P0:
                    this.Port = EPX64R.EPX64R_PORT0; // PORT0
                    break;
                case PortName.P1:
                    this.Port = EPX64R.EPX64R_PORT1; // PORT1
                    break;
                case PortName.P2:
                    this.Port = EPX64R.EPX64R_PORT2; // PORT2                   
                    break;
                case PortName.P3:
                    this.Port = EPX64R.EPX64R_PORT3; // PORT3                   
                    break;
                case PortName.P4:
                    this.Port = EPX64R.EPX64R_PORT4; // PORT4                   
                    break;
                case PortName.P5:
                    this.Port = EPX64R.EPX64R_PORT5; // PORT5                  
                    break;
                case PortName.P6:
                    this.Port = EPX64R.EPX64R_PORT6; // PORT6                   
                    break;
                case PortName.P7:
                    this.Port = EPX64R.EPX64R_PORT7; // PORT7                     
                    break;
                default:
                    return 1;   //ポート名間違えたら異常とする
            }

            // Input
            this.Status = EPX64R.EPX64R_InputPort(this.hDevice, this.Port, ref this.InputValue);
            if (Status != EPX64R.EPX64R_OK)
            {
                EPX64R.EPX64R_Close(hDevice);   // Device Close
                //MessageBox.Show("EPX64R_InputPort() Error");
                return 1;
            }

            switch (pName)
            {
                case PortName.P0:
                    this.P0InputData = this.InputValue;
                    break;
                case PortName.P1:
                    this.P1InputData = this.InputValue;
                    break;
                case PortName.P2:
                    this.P2InputData = this.InputValue;
                    break;
                case PortName.P3:
                    this.P3InputData = this.InputValue;
                    break;
                case PortName.P4:
                    this.P4InputData = this.InputValue;
                    break;
                case PortName.P5:
                    this.P5InputData = this.InputValue;
                    break;
                case PortName.P6:
                    this.P6InputData = this.InputValue;
                    break;
                case PortName.P7:
                    this.P7InputData = this.InputValue;
                    break;
            }
            return 0;
        }



        //****************************************************************************
        //メソッド名：OutByte（指定ポートにバイト単位での出力）
        //引数：引数①　ポート名（"P0"、"P1"・・・"P7"）引数②出力値（0x00～0xFF）
        //戻り値：正常０、異常１
        //****************************************************************************
        public int OutByte(PortName pName, byte Data)
        {
            int flag = 0;

            //ポートの特定と出力データバッファの更新
            switch (pName)
            {
                case PortName.P0:
                    this.Port = EPX64R_PORT0;
                    this.p0Outdata = Data;
                    break;
                case PortName.P1:
                    this.Port = EPX64R_PORT1;
                    this.p1Outdata = Data;
                    break;
                case PortName.P2:
                    this.Port = EPX64R_PORT2;
                    this.p2Outdata = Data;
                    break;
                case PortName.P3:
                    this.Port = EPX64R_PORT3;
                    this.p3Outdata = Data;
                    break;
                case PortName.P4:
                    this.Port = EPX64R_PORT4;
                    this.p4Outdata = Data;
                    break;
                case PortName.P5:
                    this.Port = EPX64R_PORT5;
                    this.p5Outdata = Data;
                    break;
                case PortName.P6:
                    this.Port = EPX64R_PORT6;
                    this.p6Outdata = Data;
                    break;
                case PortName.P7:
                    this.Port = EPX64R_PORT7;
                    this.p7Outdata = Data;
                    break;
                default:
                    return 1;//ポート名間違えたら異常とする

            }

            //データの出力
            flag = EPX64R.EPX64R_OutputPort(this.hDevice, this.Port, Data);
            if (flag != EPX64R.EPX64R_OK)
            {
                EPX64R.EPX64R_Close(this.hDevice);   // Device Close
                //MessageBox.Show("EPX64R_OutputPort() Error");
                return 1;
            }

            return 0;
        }

        //****************************************************************************
        //メソッド名：OutBit（指定ポートにビット単位での出力）
        //引数：引数①　ポート名（"P00"、"P01"・・・"P07"）引数②出力値（H：0/ L: 1）
        //戻り値：正常０、異常１
        //****************************************************************************
        public int OutBit(PortName pName, BitName bName, OutData data)
        {

            //ポートの特定
            byte PortOutData = 0;

            switch (pName)
            {
                case PortName.P0:
                    this.Port = EPX64R_PORT0;
                    PortOutData = this.p0Outdata;
                    break;
                case PortName.P1:
                    this.Port = EPX64R_PORT1;
                    PortOutData = this.p1Outdata;
                    break;
                case PortName.P2:
                    this.Port = EPX64R_PORT2;
                    PortOutData = this.p2Outdata;
                    break;
                case PortName.P3:
                    this.Port = EPX64R_PORT3;
                    PortOutData = this.p3Outdata;
                    break;
                case PortName.P4:
                    this.Port = EPX64R_PORT4;
                    PortOutData = this.p4Outdata;
                    break;
                case PortName.P5:
                    this.Port = EPX64R_PORT5;
                    PortOutData = this.p5Outdata;
                    break;
                case PortName.P6:
                    this.Port = EPX64R_PORT6;
                    PortOutData = this.p6Outdata;
                    break;
                case PortName.P7:
                    this.Port = EPX64R_PORT7;
                    PortOutData = this.p7Outdata;
                    break;
                default:
                    return 1;
            }

            //ビットの特定
            byte Temp = 0;
            int Num = 0;

            switch (bName)
            {
                case BitName.b0:
                    Num = 0; Temp = 0xFE;
                    break;
                case BitName.b1:
                    Num = 1; Temp = 0xFD;
                    break;
                case BitName.b2:
                    Num = 2; Temp = 0xFB;
                    break;
                case BitName.b3:
                    Num = 3; Temp = 0xF7;
                    break;
                case BitName.b4:
                    Num = 4; Temp = 0xEF;
                    break;
                case BitName.b5:
                    Num = 5; Temp = 0xDF;
                    break;
                case BitName.b6:
                    Num = 6; Temp = 0xBF;
                    break;
                case BitName.b7:
                    Num = 7; Temp = 0x7F;
                    break;
            }


        //データの出力
            byte Data = 0;
            switch (data)
            {
                case OutData.H:
                    Data = 1;
                    break;
                
                case OutData.L:
                    Data = 0;
                    break;
            }


            byte OutputValue = (byte)((PortOutData & Temp) | (Data << Num));//byteでｷｬｽﾄしないと怒られる
            int flag = EPX64R.EPX64R_OutputPort(this.hDevice, Port, OutputValue);
            if (flag != EPX64R.EPX64R_OK)
            {
                EPX64R.EPX64R_Close(this.hDevice);   // Device Close
                //MessageBox.Show("EPX64R_OutputPort() Error");
                return 1;
            }

            switch (pName)
            {
                case PortName.P0:
                    this.p0Outdata = OutputValue;
                    break;
                case PortName.P1:
                    this.p1Outdata = OutputValue;
                    break;
                case PortName.P2:
                    this.p2Outdata = OutputValue;
                    break;
                case PortName.P3:
                    this.p3Outdata = OutputValue;
                    break;
                case PortName.P4:
                    this.p4Outdata = OutputValue;
                    break;
                case PortName.P5:
                    this.p5Outdata = OutputValue;
                    break;
                case PortName.P6:
                    this.p6Outdata = OutputValue;
                    break;
                case PortName.P7:
                    this.p7Outdata = OutputValue;
                    break;
            }
            return 0;
        }





    }
}