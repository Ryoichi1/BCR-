using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

/// <summary>
/// EPX-64R API library definitions
/// </summary>

namespace BCR�ϓd��_�≏��R����
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
        
        //�N���X�ϐ��i�萔�j�錾 
        // Device status (Return codes) ��const �����o�[�̓N���X�ɑ����܂��i�ÓI�ϐ��Ɠ��������j
        public const int EPX64R_OK = 0;
        public const int EPX64R_INVALID_HANDLE = 1;
        public const int EPX64R_DEVICE_NOT_FOUND = 2;
        public const int EPX64R_DEVICE_NOT_OPENED = 3;
        public const int EPX64R_OTHER_ERROR = 4;
        public const int EPX64R_COMMUNICATION_ERROR = 5;
        public const int EPX64R_INVALID_PARAMETER = 6;

        // Constants�@��const �����o�[�̓N���X�ɑ����܂��i�ÓI�ϐ��Ɠ��������j
        public const byte EPX64R_PORT0 = 0x01;
        public const byte EPX64R_PORT1 = 0x02;
        public const byte EPX64R_PORT2 = 0x04;
        public const byte EPX64R_PORT3 = 0x08;
        public const byte EPX64R_PORT4 = 0x10;
        public const byte EPX64R_PORT5 = 0x20;
        public const byte EPX64R_PORT6 = 0x40;
        public const byte EPX64R_PORT7 = 0x80;


        //�񋓌^�̐錾
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

        //�C���X�^���X�ϐ��錾�i��������C���X�^���X�ŗL�̒l�j
        public System.IntPtr hDevice;   //Device Handle�i�|�C���^�ւ̃|�C���^�j
        public int Status;              // Device Status (Return Code)
        public int Number;              // Devices Number
        public byte Direction;          //�e�|�[�g ����or�o�͂̐ݒ�p�p�����[�^
        public byte Port;               //���o�͂̃|�[�g�w��p�p�����[�^
        public byte InputValue;         //�w��|�[�g����ǂݍ��񂾃f�[�^

        public byte P0InputData; //�|�[�g0�̓��̓f�[�^�o�b�t�@
        public byte P1InputData; //�|�[�g1�̓��̓f�[�^�o�b�t�@
        public byte P2InputData; //�|�[�g2�̓��̓f�[�^�o�b�t�@
        public byte P3InputData; //�|�[�g3�̓��̓f�[�^�o�b�t�@
        public byte P4InputData; //�|�[�g4�̓��̓f�[�^�o�b�t�@
        public byte P5InputData; //�|�[�g5�̓��̓f�[�^�o�b�t�@
        public byte P6InputData; //�|�[�g6�̓��̓f�[�^�o�b�t�@
        public byte P7InputData; //�|�[�g7�̓��̓f�[�^�o�b�t�@

        public byte p0Outdata; //���݂̃|�[�g�O�ɏo�͂���Ă���f�[�^�i�d�������Ă邩�`�F�b�N���邽��public�ɂ��Ă����j
        public byte p1Outdata; //���݂̃|�[�g�P�ɏo�͂���Ă���f�[�^
        public byte p2Outdata; //���݂̃|�[�g�Q�ɏo�͂���Ă���f�[�^
        public byte p3Outdata; //���݂̃|�[�g�R�ɏo�͂���Ă���f�[�^
        public byte p4Outdata; //���݂̃|�[�g�S�ɏo�͂���Ă���f�[�^
        public byte p5Outdata; //���݂̃|�[�g�T�ɏo�͂���Ă���f�[�^
        public byte p6Outdata; //���݂̃|�[�g�U�ɏo�͂���Ă���f�[�^
        public byte p7Outdata; //���݂̃|�[�g�V�ɏo�͂���Ă���f�[�^

        //�C���X�^���X�R���X�g���N�^
        public EPX64R()
        {
            //�C���X�^���X�ϐ��̏�����
            hDevice = System.IntPtr.Zero;   // Device Handle
            Status = 0; // Device Status (Return Code)
            Number = 0; // Devices Number
            Direction = 0;
            Port = 0;
            InputValue = 0;

            P0InputData = 0; //���݂̃|�[�g�O�ɏo�͂���Ă���f�[�^
            P1InputData = 0; //���݂̃|�[�g�P�ɏo�͂���Ă���f�[�^
            P2InputData = 0; //���݂̃|�[�g�Q�ɏo�͂���Ă���f�[�^
            P3InputData = 0; //���݂̃|�[�g�R�ɏo�͂���Ă���f�[�^
            P4InputData = 0; //���݂̃|�[�g�S�ɏo�͂���Ă���f�[�^
            P5InputData = 0; //���݂̃|�[�g�T�ɏo�͂���Ă���f�[�^
            P6InputData = 0; //���݂̃|�[�g�U�ɏo�͂���Ă���f�[�^
            P7InputData = 0; //���݂̃|�[�g�V�ɏo�͂���Ă���f�[�^

            p0Outdata = 0; //���݂̃|�[�g�O�ɏo�͂���Ă���f�[�^
            p1Outdata = 0; //���݂̃|�[�g�P�ɏo�͂���Ă���f�[�^
            p2Outdata = 0; //���݂̃|�[�g�Q�ɏo�͂���Ă���f�[�^
            p3Outdata = 0; //���݂̃|�[�g�R�ɏo�͂���Ă���f�[�^
            p4Outdata = 0; //���݂̃|�[�g�S�ɏo�͂���Ă���f�[�^
            p5Outdata = 0; //���݂̃|�[�g�T�ɏo�͂���Ă���f�[�^
            p6Outdata = 0; //���݂̃|�[�g�U�ɏo�͂���Ă���f�[�^
            p7Outdata = 0; //���݂̃|�[�g�V�ɏo�͂���Ă���f�[�^
        }


        //**************************************************************************
        //EPX64�̏�����
        //�����F�Ȃ�
        //�ߒl�F�Ȃ� ���ُ펞�͋����I��
        //**************************************************************************
        public void InitEpx64r(byte directionData)
        {
            try
            {

                // Device Number
                this.Status = EPX64R.EPX64R_GetNumberOfDevices(ref this.Number);
                if (this.Status != EPX64R.EPX64R_OK)
                {
                    MessageBox.Show("EPX64R_GetNumberOfDevices() Error" + "\n" + "�����I�����܂�", "IO�ް�ވُ�");
                    Environment.Exit(0);
                }
                if (this.Number == 0)
                {
                    MessageBox.Show("DeviceNumber=0" + "\n" + "�����I�����܂�", "IO�ް�ވُ�");
                    Environment.Exit(0);
                }

                // Device Open
                this.Status = EPX64R.EPX64R_Open(ref this.hDevice);
                if (this.Status != EPX64R.EPX64R_OK)
                {
                    MessageBox.Show("EPX64R_Open() Error" + "\n" + "�����I�����܂�", "IO�ް�ވُ�");
                    Environment.Exit(0);
                }

                this.Direction = directionData; // Direction 0=���� 1=�o��
                this.Status = EPX64R.EPX64R_SetDirection(this.hDevice, this.Direction);
                if (this.Status != EPX64R.EPX64R_OK)
                {
                    EPX64R.EPX64R_Close(this.hDevice);   // Device Close
                    MessageBox.Show("EPX64R_SetDirection() Error" + "\n" + "�����I�����܂�", "IO�ް�ވُ�");
                    return;
                }
            }
            catch
            {
                //EPX64R.EPX64R_Close(this.hDevice);   // Device Close
                MessageBox.Show("EPX64R_SetDirection() Error" + "\n" + "�����I�����܂�", "IO�ް�ވُ�");
                Environment.Exit(0);
                return;
            }
        }

        //****************************************************************************
        //���\�b�h���FReadInputData�i�w��|�[�g�̃f�[�^����荞�ށj
        //�����F�|�[�g���i"P0"�A"P1"�E�E�E"P7"�j
        //�߂�l�F��荞�񂾃f�[�^����O�A�ُ�P
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
                    return 1;   //�|�[�g���ԈႦ����ُ�Ƃ���
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
        //���\�b�h���FOutByte�i�w��|�[�g�Ƀo�C�g�P�ʂł̏o�́j
        //�����F�����@�@�|�[�g���i"P0"�A"P1"�E�E�E"P7"�j�����A�o�͒l�i0x00�`0xFF�j
        //�߂�l�F����O�A�ُ�P
        //****************************************************************************
        public int OutByte(PortName pName, byte Data)
        {
            int flag = 0;

            //�|�[�g�̓���Əo�̓f�[�^�o�b�t�@�̍X�V
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
                    return 1;//�|�[�g���ԈႦ����ُ�Ƃ���

            }

            //�f�[�^�̏o��
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
        //���\�b�h���FOutBit�i�w��|�[�g�Ƀr�b�g�P�ʂł̏o�́j
        //�����F�����@�@�|�[�g���i"P00"�A"P01"�E�E�E"P07"�j�����A�o�͒l�iH�F0/ L: 1�j
        //�߂�l�F����O�A�ُ�P
        //****************************************************************************
        public int OutBit(PortName pName, BitName bName, OutData data)
        {

            //�|�[�g�̓���
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

            //�r�b�g�̓���
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


        //�f�[�^�̏o��
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


            byte OutputValue = (byte)((PortOutData & Temp) | (Data << Num));//byte�ŷ��Ă��Ȃ��Ɠ{����
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