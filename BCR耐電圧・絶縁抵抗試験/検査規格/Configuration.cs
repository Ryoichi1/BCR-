using System.Collections.Generic;

namespace BCR耐電圧_絶縁抵抗試験
{
    public class Configuration
    {
        public string 日付 { get; set; }

        public int TodayOkCount { get; set; }
        public int TodayNgCount { get; set; }

        public string PathTheme { get; set; }
        public double OpacityTheme { get; set; }
        public List<string> 作業者リスト { get; set; }
    }
}
