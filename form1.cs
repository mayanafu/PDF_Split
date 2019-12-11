using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using iTextSharp.text.pdf;
using iTextSharp.text;


//■処理概要■
//PDFを分割出力する。
//重複しているしおりのデータは出力しない。

namespace PDF_Split
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            ProgressBar1.Visible = false;

            txtOutput.Text = "C:\test";


            cmbYEAR.Items.Add(DateTime.Now.AddYears(-1).ToString("yyyy"));
            cmbYEAR.Items.Add(DateTime.Now.ToString("yyyy"));
            cmbYEAR.Items.Add(DateTime.Now.AddYears(1).ToString("yyyy"));
            cmbYEAR.Parent = this;

            if (DateTime.Now.ToString("yyyy") == DateTime.Now.AddMonths(7).ToString("yyyy"))
            {
                cmbYEAR.SelectedIndex = 1;
            }
            else
            {
                cmbYEAR.SelectedIndex = 2;
            }


            for (int i = 1; i < 13; ++i)
            {
                cmbMONTH.Items.Add(i);
            }

            cmbMONTH.Parent = this;
            cmbMONTH.SelectedIndex = int.Parse(DateTime.Now.AddMonths(7).ToString("MM")) - 1;


        }

        private void btnInput_Click(object sender, EventArgs e)
        {

            OpenFileDialog dr = new OpenFileDialog();
            //はじめに表示されるフォルダを指定する
            dr.InitialDirectory = "C:\";
            dr.Filter = "PDFファイル (*.pdf)|*.pdf";

            if (dr.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtInput.Text = dr.FileName;
            }

        }

        private void btnOutput_Click(object sender, EventArgs e)
        {


            //FolderBrowserDialogクラスのインスタンスを作成
            FolderBrowserDialog fbd = new FolderBrowserDialog();

            //上部に表示する説明テキストを指定する
            fbd.Description = "フォルダを指定してください。";
            //ルートフォルダを指定する
            //デフォルトでDesktop
            fbd.RootFolder = Environment.SpecialFolder.Desktop;
            //最初に選択するフォルダを指定する
            //RootFolder以下にあるフォルダである必要がある
            fbd.SelectedPath = @"C:\";
            //ユーザーが新しいフォルダを作成できるようにする
            //デフォルトでTrue
            fbd.ShowNewFolderButton = true;

            //ダイアログを表示する
            if (fbd.ShowDialog(this) == DialogResult.OK)
            {
                //選択されたフォルダパスをテキストボックスに表示
                txtOutput.Text = fbd.SelectedPath;
            }


        }


        private void btnPDF_Create_Click(object sender, EventArgs e)
        {


            if (txtInput.Text == "")
            {
                MessageBox.Show("参照元PDFを選択して下さい。");
                btnInput.Focus();
                return;
            }

            if (txtOutput.Text == "")
            {
                MessageBox.Show("保存先を選択して下さい。");
                btnOutput.Focus();
                return;
            }

            DialogResult dr = MessageBox.Show("処理を開始します。よろしいですか？", "確認", MessageBoxButtons.YesNo);

            if (dr == System.Windows.Forms.DialogResult.Yes)
            {
            }
            else if (dr == System.Windows.Forms.DialogResult.No)
            {
                return;
            }
            else
            {
                return;
            }



            //コントロールを初期化する
            ProgressBar1.Minimum = 0;
            ProgressBar1.Value = 0;
            Label1.Text = "";
            //Label1を再描画する
            Label1.Update();


            // 参照元PDFファイルのオープン
            PdfReader rd = new PdfReader(txtInput.Text);

            //出力先pdf用
            PdfCopyFields oPdfcopy = null;
            Document oDocument = null;

            // ページ数を取得
            int iEND = rd.NumberOfPages;


            // しおりの情報を取得
            IList<Dictionary<string, object>> obookmarkList = SimpleBookmark.GetBookmark(rd);


            // ＰＤＦを閉じる
            rd.Close();


            int i; // カウント用変数
            string[] sTITLE = new string[obookmarkList.Count]; // しおりの件数をセット(しおり名)
            int[] iSectionStart = new int[obookmarkList.Count]; // しおりの件数をセット(しおりの開始位置)
            int[] iSectionEnd = new int[obookmarkList.Count]; // しおりの件数をセット(しおりの終了位置)
            ProgressBar1.Maximum =   obookmarkList.Count; // プログレスバーの最大値設定
            string[] sMONTH = new string[13] { "","Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" }; //英字変換用



            try
            {


                for (i = 0; i < obookmarkList.Count; ++i) // iをしおりの件数分、1ずつ増やして繰り返し
                {

                    // しおりの名前
                    string[] sSHIORI = ((obookmarkList[i]["Title"]).ToString()).Split('@');
                    sTITLE[i] = sSHIORI[2].ToString().Replace("CSTCNTR:", "");


                    // 開始ページ数
                    string[] sStartPage = ((obookmarkList[i]["Page"]).ToString()).Split(' ');
                    iSectionStart[i] = Convert.ToInt32(sStartPage[0]);



                    // 終了ページ数
                    int w_sEndPage;


                    int j = i;

                    do
                    {
                        if (j + 1 == obookmarkList.Count)
                        {
                            w_sEndPage = iEND;
                        }
                        else
                        {
                            string[] sEndPage = ((obookmarkList[j + 1]["Page"]).ToString()).Split(' ');  //次のしおりの開始ページを取得
                            w_sEndPage = Convert.ToInt32(sEndPage[0]);
                            j = j + 1;
                        }


                    }
                    while (iSectionStart[i] == w_sEndPage && i < obookmarkList.Count - 1 - 1); //現在のしおりの開始ページと次のしおりの開始ページが同じ時(重複時)、
                                                                                               //その次のしおりの開始ページを読み込む(ループ処理を続ける)
                                                                                               //※しかし、最終の重複時は処理を抜ける。

                    if (i + 1 == obookmarkList.Count) //最終データ

                    { iSectionEnd[i] = iEND; }

                    else if (iSectionStart[i] == w_sEndPage && i == obookmarkList.Count - 1 - 1) //最終の一つ前のデータ(しおり重複時)
                                                                                                 //※しおりカウント(obookmarkList.Count)は1から始まっているので、処理変数iの位置と合わせるために-1、
                                                                                                 //  更に最終の一つ前のデータなので、更に-1

                    { iSectionEnd[i] = iEND; }

                    else  //そのほかのデータ

                    { iSectionEnd[i] = w_sEndPage - 1; }


                }


                oDocument = new Document(rd.GetPageSizeWithRotation(1));
                oDocument.Open();

                string SHIORI_NAME = "";
                int w_SectionEnd = 0;
                int w_COUNT = 1;
                int k = 0;

                ProgressBar1.Visible = true;

                for (i = 0; i < obookmarkList.Count; ++i) // iをしおりの件数分、1ずつ増やして繰り返し
                {

                    ProgressBar1.Value = i;


                    if (sTITLE[i] == SHIORI_NAME)


                    { } //前回処理した時としおり名が一致しているとき、処理を行わない。

                    else
 
                    {
                        oPdfcopy = new PdfCopyFields(new FileStream(txtOutput.Text + "\\" + "MAN " + sMONTH[cmbMONTH.SelectedIndex + 1] + " " + cmbYEAR.SelectedItem + "_" + w_COUNT  + ".pdf", FileMode.Create));


                        k = 0;


                        //同しおり名の最後のページを記憶
                        do
                        {
                            w_SectionEnd = iSectionEnd[i + k];
                            k = k + 1;
                            if (i + k >= obookmarkList.Count - 1) { break;}
                        }
                        while (sTITLE[i] == sTITLE[i + k] );

                        
                        rd = new PdfReader(txtInput.Text);

                        rd.SelectPages(iSectionStart[i].ToString() + "-" + w_SectionEnd.ToString()); //現在のしおりの該当ページ部分を選択

                        oPdfcopy.AddDocument(rd); //PDFアウトプット
                        oPdfcopy.Close(); //pdf保存
                        rd.Close();

                        w_COUNT = w_COUNT + 1;


                        SHIORI_NAME = sTITLE[i]; //処理したしおり名を記憶しておく


                        Label1.Text = "PDF分割完了：" + "MAN " + sMONTH[cmbMONTH.SelectedIndex + 1] + " " + cmbYEAR.SelectedItem + "_" + w_COUNT + ".pdf";
                        Label1.Update();
                    }
                }

                ProgressBar1.Value = 0;
                ProgressBar1.Visible = false;
                Label1.Text = "処理完了：" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");


            }


            catch (Exception ex)
            {
                MessageBox.Show("エラー発生： " + ex.ToString());
            }


            finally
            {
                if (rd        != null) { rd.Close(); }
                if (oPdfcopy  != null) { oPdfcopy.Close(); }
                if (oDocument != null) { oDocument.Close(); }
            }

        
    } 
 

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}


