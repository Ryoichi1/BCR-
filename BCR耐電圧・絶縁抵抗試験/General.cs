using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.IO;
using System.Media;

namespace BCR耐電圧_絶縁抵抗試験
{
    public static class General
    {
        public static SoundPlayer player = null;


        //**************************************************************************
        //WAVEファイル関連
        //引数：
        //戻値：
        //**************************************************************************

        //WAVEファイルを再生する（再生中は他の割り込み処理を受け付けない）
        public static void PlaySound(string waveFile)
        {
            //再生されているときは止める
            if (player != null)
                player.Stop();

            //waveファイルを読み込む
            player = new System.Media.SoundPlayer(waveFile);
            //最後まで再生し終えるまで待機する
            //player.Play();
            player.PlaySync();
        }
        //WAVEファイルを再生する（非同期で再生）
        public static void PlaySound2(string waveFile)
        {
            //再生されているときは止める
            if (player != null)
                player.Stop();

            //waveファイルを読み込む
            player = new System.Media.SoundPlayer(waveFile);
            //最後まで再生し終えるまで待機する
            player.Play();
        }
        //WAVEファイルをループ再生する
        public static void PlayLoopSound(string waveFile)
        {
            //再生されているときは止める
            if (player != null)
                player.Stop();

            //waveファイルを読み込む
            player = new System.Media.SoundPlayer(waveFile);
            //次のようにすると、ループ再生される
            player.PlayLooping();
        }
        //再生されているWAVEファイルを止める
        public static void StopSound()
        {
            if (player != null)
            {
                player.Stop();
                player.Dispose();
                player = null;
            }
        }




        //**************************************************************************
        //ウェイト
        //引数：指定時間（ｍｓｅｃ）
        //戻値：なし
        //**************************************************************************
        public static void Wait(int time)
        {
            foreach (int i in Enumerable.Range(0, time / 10)) //for (int i = 0; i < (time / 10); i++)
            {
                Thread.Sleep(10);
                Application.DoEvents();
            }
        }

        //**************************************************************************
        //Openoffice（calc）の強制終了（プロセスの強制終了）
        //引数：
        //戻値：
        //**************************************************************************
        public static void KillOpenOffice()
        {

            try
            {
                //ローカルコンピュータ上で実行されているすべてのプロセスを取得
                System.Diagnostics.Process[] ps = System.Diagnostics.Process.GetProcesses();


                //配列から1つずつ取り出す
                foreach (System.Diagnostics.Process p in ps)
                {
                    //プロセスの中に"soffice.bin"があれば強制的にプロセスを終了する
                    if (p.ProcessName == "soffice.bin") p.Kill();
                }
            }
            catch
            {
                MessageBox.Show("OpenOfficeのプロセスが残ったままです！\r\nタスクマネージャからプロセスを終了してください");
            }
            
        }


    }
}
