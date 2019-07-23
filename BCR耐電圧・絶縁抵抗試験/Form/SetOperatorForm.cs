using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace BCR耐電圧_絶縁抵抗試験
{
    public partial class SetOperatorForm : Form
    {
        
        public SetOperatorForm()
        {
            InitializeComponent();
        }

        //**************************************************************************
        //フォームロードイベント
        //引数：
        //戻値：
        //**************************************************************************
        private void SetOperatorForm_Load(object sender, EventArgs e)
        {
            foreach (string name in State.作業者リスト)
            {
                listBoxOperatorName.Items.Add(name);
            }
            
        }

        //**************************************************************************
        //キャンセルボタンを押したときの処理
        //引数：
        //戻値：
        //**************************************************************************
        private void buttonCancel_Click(object sender, EventArgs e)
        {
            buttonOk.Enabled = false;
            buttonCancel.Enabled = false;
            this.Close();
        }

        //**************************************************************************
        //ＯＫボタンを押したときの処理
        //引数：
        //戻値：
        //**************************************************************************
        private void buttonOk_Click(object sender, EventArgs e)
        {
            buttonOk.Enabled = false;
            buttonCancel.Enabled = false;
            
            //パラメータファイル　作業者一覧を更新する
                       
            var name = new List<string>();
            foreach(string item in listBoxOperatorName.Items)
            {
                name.Add(item);
            }

            foreach (int i in Enumerable.Range(0, 10 - name.Count))
            {
                name.Add("予約");
            }

            //パラメータファイルの作業者名一覧を更新
            if (!SetOperatorName(name))
            {
                MessageBox.Show("パラメータファイルの更新に失敗しました");
                return;
            }
            
            this.Close();
        }

        //**************************************************************************
        //作業者一覧の更新
        //引数：
        //戻値：
        //**************************************************************************
        private bool SetOperatorName(List<string> name)
        {
            OpenOffice calc = new OpenOffice();
            //parameterファイルを開く
            calc.OpenFile(Constants.ParameterFilePath);


            // sheetを取得           
            calc.SelectSheet("OperatorName");


            //作業者一覧の更新

            int i = 0;
            //行＝ROW 列＝COLUMN 
            foreach(string item in name)
            {
                calc.cell = calc.sheet.getCellByPosition(1, 1 + i);
                calc.cell.setFormula(item);
                i++;
            }

            // Calcファイルを保存して閉じる
            if (!calc.SaveFile()) return false;
            
            return true;
        }

        //**************************************************************************
        //追加ボタンを押したときの処理
        //引数：
        //戻値：
        //**************************************************************************
        private void buttonAdd_Click(object sender, EventArgs e)
        {
            if (textBoxNewName.Text == "")
            {
                return;
            }

            if (listBoxOperatorName.Items.Count > 9)
            {
                MessageBox.Show("作業者登録は10名までです");
                textBoxNewName.Text = "";
                return;

            }
            listBoxOperatorName.Items.Add(textBoxNewName.Text);
            textBoxNewName.Text = "";

        }

        //**************************************************************************
        //削除ボタンを押したときの処理
        //引数：
        //戻値：
        //**************************************************************************
        private void buttonDelete_Click(object sender, EventArgs e)
        {
            if (listBoxOperatorName.SelectedIndex < 0)
            {
                MessageBox.Show("削除する名前が選択されていません");
                return;
            }

            listBoxOperatorName.Items.RemoveAt(listBoxOperatorName.SelectedIndex);
        
        }
    }
}
