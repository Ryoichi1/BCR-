using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BCR耐電圧_絶縁抵抗試験
{
    public static class Flags
    {
        public static Action<bool> 日常点検ラベルのセット;


        //各種フラグ
        public  static bool FlagOpeName; //作業者名が正しく選択されていればTrue
        public  static bool FlagSerial;　//シリアルが正しく入力されていればTrue
        public  static bool FlagBcr;

        private static bool FlagDailyCheck;//日常点検が実施されていればTrue



        
        public  static bool FlagTest;
        public  static bool FlagCount;
        
        //プロパティ
        public static bool 日常点検
        {
            get
            {
                return FlagDailyCheck;
            }
            set
            {
                FlagDailyCheck = value;
                日常点検ラベルのセット(value);
            }
        }



    }
}
