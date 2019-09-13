namespace BCR耐電圧_絶縁抵抗試験 
{
    public class GeneralTimer
    {
        //.Net FrameWork 4.5　にて下記プログラム作成しました。
        //Taskで実行するとtimerはtimer_Tickにたどり着かず、
        //そのまま呼び出すとtimer_Tickに行きます。
        //両者の違いは何故発生するのでしょうか？
        //よろしくお願いします。

        ///System.Windows.Forms.Timerを使用していたことが原因でした。
        //System.Timers.Timerを使用することでTaskて使用してもTick (Elapsed)することがわかりました。




        public bool FlagTimeout { get; private set; }
        public int Time { get; set; }

        public bool IsRunning { get; set; }


        private System.Timers.Timer TmTimeOut;

        public GeneralTimer(int time_msec)
        {
            TmTimeOut = new System.Timers.Timer();
            this.Time = time_msec;
            TmTimeOut.Elapsed += (sender, e) =>
            {
                TmTimeOut.Stop();
                FlagTimeout = true;
            };
        }

        public void Start(int NewTime_msec = 0)
        {
            if (NewTime_msec == 0)
            {
                TmTimeOut.Interval = Time;
            }
            else
            {
                TmTimeOut.Interval = NewTime_msec;
            }

            TmTimeOut.Stop();
            FlagTimeout = false;
            TmTimeOut.Start();
            IsRunning = true;
        }

        public void stop()
        {
            TmTimeOut.Stop();
            IsRunning = false;
        }
    }
}
