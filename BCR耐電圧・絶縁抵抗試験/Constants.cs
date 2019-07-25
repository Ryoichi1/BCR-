using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BCR耐電圧_絶縁抵抗試験
{
    public static class Constants
    {
        //定数の宣言

        //検査ソフトVer
        //public const string CheckerSoftVer = "1.00";//新規作成
        //public const string CheckerSoftVer = "2.00";//変更日時（2014.3.25）
                                                    //①クラス構成をステート集約型に変更
                                                    //②渡辺さんからの指示で試験ログデータの保存ファイル名を変更
        //public const string CheckerSoftVer = "3.00";//変更日時（2014.7.4）AUR機種追加
        public const string CheckerSoftVer = "3.10";//変更日時（2019.7.24）BC-R 対アース間の耐電圧検査追加対応

        //ファイルパスいろいろ
        public const string ParameterFilePath = @"C:\BCR耐電圧・絶縁抵抗試験\parameter.ods";

        public const string Pc3LogFileFolder = @"\\Y01130003\bc-r_data\BCR_PC3\Log\";
        public const string Pc4LogFileFolder = @"\\Y01130003\bc-r_data\BCR_PC4\Log\";
        public const string Pc7LogFileFolder = @"\\Y01130003\bc-r_data\BCR_PC7\Log_AUR\";//AUR誤配線チェックの結果があります

        public const string 日常点検FolderPath = @"C:\BCR耐電圧・絶縁抵抗試験\日常点検\";

        public const string 検査ﾃﾞｰﾀBcrFolderPath = @"C:\BCR耐電圧・絶縁抵抗試験\検査データ\";
        
        public const string SoundSet            = @"C:\BCR耐電圧・絶縁抵抗試験\Wav\Set.wav";
        public const string SoundWarning        = @"C:\BCR耐電圧・絶縁抵抗試験\Wav\Warning.wav";
        public const string SoundName           = @"C:\BCR耐電圧・絶縁抵抗試験\Wav\Operator.wav";
        public const string SoundDailyCheck     = @"C:\BCR耐電圧・絶縁抵抗試験\Wav\DailyCheck.wav";
        public const string SoundPass           = @"C:\BCR耐電圧・絶縁抵抗試験\Wav\Pass.wav";
        public const string SoundFail           = @"C:\BCR耐電圧・絶縁抵抗試験\Wav\Fail.wav";

        public const string SetBcrPath              = @"C:\BCR耐電圧・絶縁抵抗試験\SetBcr.bmp";
        public const string SetAurPath              = @"C:\BCR耐電圧・絶縁抵抗試験\SetAur.bmp";

        //作業者へのメッセージ
        public const string MessOperator        = "作業者名を選択してください";
        public const string MessDailyCheck      = "日常点検を実施してください";
        public const string MessSet             = "製品をセットして、ＱＲコードを読み取ってください";
        public const string MessRemove          = "製品を取り外してください";
        public const string MessWait            = "検査中！　治具から離れてお待ちください・・・";
        public const string MessWarning         = "治具から離れてください！　高電圧印加まで ";




    }
}
