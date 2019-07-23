using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using uno.util;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.sheet;
using unoidl.com.sun.star.table;
using unoidl.com.sun.star.uno;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.util;

namespace BCR耐電圧_絶縁抵抗試験
{

    public class OpenOffice
    {

        //インスタンスメンバ
        public XSpreadsheets sheets;
        public XSpreadsheetDocument doc;
        public Uri uriCalcFile;
        public XComponentContext localContext;
        public XMultiServiceFactory multiServiceFactory;
        public XComponentLoader loader;
        public PropertyValue[] calcProp;
        public XCloseable xCloseable;
        public XStorable xstorable;
        public PropertyValue[] storeProps;
        public XSpreadsheet sheet;
        public XCell cell;

        //ファイルを開く
        public void OpenFile(string filePath)
        {
            //parameterファイルを開く
            string path = filePath;

            // pathをURIに変換
            Uri.TryCreate(path, UriKind.Absolute, out uriCalcFile);

            // OpenOfficeファイルを開くためのおまじない
            // コンポーネントコンテキストの取得
            localContext = Bootstrap.bootstrap();
            // サービスマネージャーの取得
            multiServiceFactory = (XMultiServiceFactory)localContext.getServiceManager();
            // コンポーネントローダーの取得
            loader = (XComponentLoader)multiServiceFactory.createInstance("com.sun.star.frame.Desktop");

            // 非表示で実行するためのプロパティ指定
            calcProp = new PropertyValue[1];
            calcProp[0] = new PropertyValue();
            calcProp[0].Name = "Hidden";
            calcProp[0].Value = new uno.Any(true);

            // Calcファイルを開く
            doc = (XSpreadsheetDocument)loader.loadComponentFromURL(uriCalcFile.ToString(), "_blank", 0, calcProp);
            // シート群を取得
            sheets = doc.getSheets();

        }

        //シートの洗濯
        public void SelectSheet(string sheetName)
        {
            // sheetを取得　"Spec"
            sheet = (XSpreadsheet)sheets.getByName(sheetName).Value;
        }

        //ファイルを閉じる
        public void CloseFile()
        {
            // Calcファイルを閉じる
            if (doc != null)
            {
                xCloseable = (XCloseable)doc;
                xCloseable.close(true);
                doc = null;
            }
        }

        //ファイルを上書き保存して閉じる
        public bool SaveFile()
        {
            // Calcファイルを保存
            xstorable = (XStorable)doc;
            storeProps = new PropertyValue[1];
            storeProps[0] = new PropertyValue();
            storeProps[0].Name = "Overwrite";     // 上書き
            storeProps[0].Value = new uno.Any(true);

            try
            {
                xstorable.storeAsURL(uriCalcFile.ToString(), storeProps);    // 保存
                CloseFile();
                return true;
            }
            catch (unoidl.com.sun.star.uno.Exception)
            {
                return false;
            }
        }


    }

}