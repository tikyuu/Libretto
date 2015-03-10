using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using System.Diagnostics;


using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace LibrettoCreateTool
{
    public partial class Form1 : Form
    {
        private List<string> load_files_ = new List<string>();              //読み込みファイル
        private List<global.VoiceData> voice_data_ = new List<global.VoiceData>();        //ボイスデータ
        private List<string> error_messages_ = new List<string>();          //エラーメッセージ
        private string log_path_ = System.Environment.CurrentDirectory + "\\past_log.txt";
        private global.HeadFootPos[] head_foot_pos_ = new global.HeadFootPos[5];
        private global.OutputFlags out_flg_ = new global.OutputFlags(false, false, false, false, false);

        #region 共通データ群
        
        /// <summary>
        /// ボイスデータ、エラーメッセージ、エラーカウントを初期化します。
        /// </summary>
        private void AllClear()
        {
            voice_data_.Clear();
            error_messages_.Clear();
        }
        private void GridClear()
        {
            //int size = dataGridView1.Rows.Count - 1;        //そのまま指定するとRowCountが減っていって全削除できないため
            int size = dataGridView1.Rows.Count;
            for (int i = 0; i < size; ++i)
            {
                dataGridView1.Rows.RemoveAt(0);
            }
        }
        /// <summary>
        /// エラーメッセージを追加します
        /// </summary>
        /// <param name="error"></param>
        private void ErrorAdd(string error)
        {
            error_messages_.Add(error);
        }
        /// <summary>
        /// エラーチェックを行います
        /// </summary>
        /// <returns></returns>
        private bool ErrorCheck()
        {
            if (error_messages_.Count != 0)
            {
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                foreach (var error_msg in error_messages_)
                {
                    sb.AppendLine(error_msg);
                }
                MessageBox.Show(sb.ToString(), "エラーッス!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AllClear();
                return false;
            }
            return true;
        }
        private bool VoiceCheck_()
        {
            if (dataGridView1.Rows.Count == 0)
            {
                ErrorAdd("「ボイス設定」がされてないですよ。");
                return false;
            }
            return true;
        }
        private void VoiceCellCheck_(int x)
        {
            HashSet<string> hash = new HashSet<string>();

            for (int i = 0; i < dataGridView1.Rows.Count; ++i)  //dataGridView1.Rows.Count - 1
            {
                if (dataGridView1[x, i].Value == null)
                {
                    ErrorAdd("「ボイス設定」 : [ " + i.ToString() +  " , " + x.ToString() + " ] : 入力してください。");
                }
                //else if (!hash.Add(dataGridView1[x, i].Value.ToString()))
                //{
                //    ErrorAdd("「ボイス設定」 : " + dataGridView1[x, i].Value + " : 既に存在しています。");
                //}
            }
        }
        private void VoiceCellValueCheck_(int x)
        {
            HashSet<string> hash = new HashSet<string>();
            int error_index = 0;
            try
            {
                for (int i = 0; i < dataGridView1.Rows.Count; ++i)
                {
                    error_index = i;
                    //連番チェック
                    int.Parse(dataGridView1[x, i].Value.ToString());
                }
            }
            catch
            {
                object value = dataGridView1[x, error_index].Value;
                if (value == null)
                {
                    ErrorAdd("「ボイス設定」 : [ " + error_index.ToString() + " , " + x.ToString() + " ] : 数値を入力してください。");
                }
                else
                {
                    ErrorAdd("「ボイス設定 」: " + value.ToString() +" : 数値を入力してください。");
                }
                int serial = ("00000").Length;
                var s = String.Format("{0:D" + serial + "}", 0);
                dataGridView1[x, error_index].Value = s;
            }
        }
        private void KeyCheck_()
        {
            char[] targets = { ' ', ':', ';', '/', '|', ',', '*', '?', '<', '>', '\n', '\t', '\\' };
            //dataGridView1.Rows[1]
            for (int i = 0; i < dataGridView1.Rows.Count; )
            {
                foreach (var t in targets)
                {
                    if (dataGridView1.Rows[i].Cells[1].Value == null)
                    {
                        dataGridView1.Rows.RemoveAt(i);
                        goto hoge;
                    }
                    if (dataGridView1.Rows[i].Cells[1].Value.ToString().Contains(t))
                    {
                        ErrorAdd(dataGridView1.Rows[i].Cells[1].Value.ToString() + " : キャラ名に使用できない文字が含まれています");
                        break;
                    }
                }
                ++i;
                hoge: ;
            }
        }
        private void FileNameCheck()
        {
	        Hashtable voice_hash = new Hashtable();

	        for(int i = 0; i < dataGridView1.Rows.Count; ++i)
	        {
		        // key = タグ(A001)
		        string key = (string)dataGridView1[2, i].Value;
		        // value = キャラ名(めんま)
		        string value = (string)dataGridView1[1, i].Value;
		        // カブっていないキー(key)のみ取得
		        if(!voice_hash.ContainsKey(key))
			        voice_hash[key] = value;
	        }

	        // キャラ名からファイル名に使用できない文字がある場合エラーを吐き出す。
	        foreach(var value in voice_hash.Values)
	        {
		        // ファイル名に使用できない文字を取得
		        char[] invalidChars = System.IO.Path.GetInvalidFileNameChars();
                if (((string)value).IndexOfAny(invalidChars) < 0)
                {
                }
                else
                {
                    ErrorAdd(value + ": ファイル名に使用できない文字がふくまれています。");
                }
	        }
        }
        private void SetKey_(string comment)
        {
            if (comment == "") ErrorAdd("「詳細設定」のコメントを選択してください。");
            if (load_files_.Count == 0) ErrorAdd("転送元「SrcDir」ファイルがないですよ。");
            if (error_messages_.Count != 0) return;

            var key_hash = new HashSet<string>();

            // ファイル読み込み、キャラ名取得
            foreach (var file in load_files_)
            {
                // ファイルチェック
                if (!System.IO.File.Exists(file))
                {
                    ErrorAdd(file + " : が存在しません。");
                    return;
                }

                var ss = System.IO.File.ReadAllLines(file, System.Text.Encoding.Default);
                foreach (var s in ss)
                {
                    // コメントを省く
                    int comment_pos = s.IndexOf(comment);
                    if (comment_pos == 0) continue;

                    // セリフ以外を省く
                    int serif_pos = s.IndexOf("「");
                    if (serif_pos == -1) continue;

                    // プリプロセッサだった場合、(#! or #?)0000_000000の13文字から～『「』までの間のキャラ名を取得する
                    int pri_pos = s.IndexOf("#");
                    if (pri_pos != -1)
                    {
                        key_hash.Add(s.Substring(13, serif_pos - 13));
                        continue;
                    }

                    key_hash.Add(s.Substring(0, serif_pos));
                }
            }
            //SortedDictionary<string, 
            SortedSet<string> sort_hash = new SortedSet<string>();
            foreach (var key in key_hash) sort_hash.Add(key);
            HashSet<string> new_key_hash = new HashSet<string>();
            foreach (var key in sort_hash) new_key_hash.Add(key);
            //KeyhashCheck(ref key_hash, 1);
            GridViewAdd(ref new_key_hash);
        }
        void GridViewAdd(ref HashSet<string> key_hash)
        {
            char c = 'A';
            int cnt = 0;
            int max = dataGridView1.Rows.Count; //増量していくのはカウントしない
            HashSet<string> now_hash = new HashSet<string>();

            //0なら追加して終了
            if (max == 0)
            {
                foreach (var key in key_hash)
                {
                    dataGridView1.Rows.Add("delete", key, c + (++cnt).ToString(), "00000", false);
                }
                return;
            }

            for (int i = 0; i < max; ++i)
            {
                if (dataGridView1[1, i].Value == null) continue;
                now_hash.Add(dataGridView1[1, i].Value.ToString());
            }
            foreach (var add_key in key_hash)
            {
                if (!now_hash.Contains(add_key))
                {
                    dataGridView1.Rows.Add("delete", add_key, c + (++cnt).ToString(), "00000", false);
                }
            }
            foreach (var now_key in now_hash)
            {
                if (!key_hash.Contains(now_key))
                {
                    ErrorAdd("【" + now_key + "】 は存在していませんが？");
                }
            }
        }
        // データグリッドのチェックボックスの更新
        private void dataGridView1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }
        #endregion
        #region フォーム１
        /// <summary>
        /// コンストラクタ
        /// </summary>
        public Form1()
        {
            InitializeComponent();
            //comment_select.Items.Add("#");
            comment_select.Items.Add("//");
            comment_select.Text = "//";
            DetailInit();
            DetailSet();

            checkBox1.Checked = true;
            checkBox2.Checked = true;
            checkBox3.Checked = true;
            checkBox5.Checked = true;

            //バージョンver
            version_text.Text = "2.0";
        }
        private void Form1_Load(object sender, EventArgs e){}
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            var ret = MessageBox.Show("データを保存しますか？", "保存", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (ret == DialogResult.Yes)
            {
                saveFileDialog1.InitialDirectory = System.Environment.CurrentDirectory;

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    //保存
                    WriteXml(saveFileDialog1.FileName);
                }
            }
        }
        #endregion
        #region スタートボタン
        Libretto libretto = new Libretto();
        global.ExcelData excel_data_;
        string dst_text_ = "";

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            libretto.Start(load_files_, dst_text_, voice_data_, excel_data_, head_foot_pos_, out_flg_, backgroundWorker1);
        }
        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

            progressBar1.Value += e.ProgressPercentage;
            string s = "現在" + progressBar1.Value + "%です。";
            string mes = (string)e.UserState;


            if (mes != null)
            {
                if (mes.IndexOf("終了") != -1)
                {
                    HistoryList.Items.RemoveAt(HistoryList.Items.Count - 1);
                    HistoryList.Items.RemoveAt(HistoryList.Items.Count - 1);
                }
                HistoryList.Items.Add(mes);
            }

            if (e.ProgressPercentage != 0)
                HistoryList.Items.Add(s);
            // 自動スクロール処理
            HistoryList.TopIndex = HistoryList.Items.Count - 1;
            
        }
        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar1.Value = 100;
            StartButton.Enabled = true;
            HistoryList.Items.Add("終了");
            // 自動スクロール処理
            HistoryList.TopIndex = HistoryList.Items.Count - 1;

            PlaySound();
            MessageBox.Show("終了");
            //StopSound();
            progressBar1.Value = 0;
        }
        /// <summary>
        /// スタートボタンクリック
        /// </summary>
        private void StartButton_Click(object sender, EventArgs e)
        {
            List<String> list_data = new List<String>();
            if (comment_select.Text == "") ErrorAdd("「詳細設定」のコメントを選択してください。");
            if (title_text.Text == "") ErrorAdd("「詳細設定」のタイトル名を入力してください。");
            if (load_files_.Count == 0) ErrorAdd("転送元「SrcDir」ファイルがないですよ。");
            if (DstBox.Text == "") ErrorAdd("転送先「DstDir」フォルダがないですよ。");

            VoiceCellCheck_(1);     //キャラ名チェック
            VoiceCellCheck_(2);     //ラベル名チェック
            VoiceCellValueCheck_(3);//連番チェック

            FileNameCheck();

            VoiceCheck_();          //キャラ名禁止文字確認
            KeyCheck_();            //キャラデータ存在確認
            if (!ErrorCheck()) return;
            voice_data_.Clear();
            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                try
                {
                    int serial = int.Parse(dataGridView1[3, i].Value.ToString());
                    int digit = dataGridView1[3, i].Value.ToString().Length;
                    bool serial_lib = Convert.ToBoolean(dataGridView1[4, i].Value);
                    //bool serial_lib = dataGridView1[4, i].Value == "True" ? true : false;
                    voice_data_.Add(new global.VoiceData(
                        dataGridView1[1, i].Value.ToString(),
                        dataGridView1[2, i].Value.ToString(),
                        serial,
                        digit,
                        voice_top_text.Text,
                        voice_middle_text.Text,
                        voice_bottom_text.Text,
                        serial_lib));
                }
                catch
                {
                    ErrorAdd((i + 1).ToString() + "行目 : なんかエラーでてますよ");
                }
            }

            // エラーチェック
            if (!ErrorCheck()) return;


            //=============start!==============
            //StartButton.Enabled = false;
            int data = int.Parse(logo_select.SelectedIndex.ToString());
            var list = new[]
            {
                new { Pos = int.Parse(logo_select.SelectedIndex.ToString()), Text = logo_text.Text },
                new { Pos = int.Parse(page_select.SelectedIndex.ToString()), Text = "" },
                new { Pos = int.Parse(file_select.SelectedIndex.ToString()), Text = "" },
                new { Pos = int.Parse(libretto_select.SelectedIndex.ToString()), Text = "" },
                new { Pos = int.Parse(date_select.SelectedIndex.ToString()), Text = "" },
            };
            for (int i = 0; i < head_foot_pos_.Length; i++)
            {
                head_foot_pos_[i] = SetHeadFoot(list[i].Pos, list[i].Text);
            }

            excel_data_ = new global.ExcelData(100, 10, comment_select.Text, title_text.Text,
                        check_voice_bold.Checked, color_text.ForeColor, color_text2.ForeColor,
                        font_select.Text, version_text.Text, date_text.Text,
                        convert_comment.Checked, check_add_star.Checked, check_visible.Checked);
            dst_text_ = DstBox.Text;


            StartButton.Enabled = false;
            // バックグランド処理開始
            Console.WriteLine("処理開始");
            backgroundWorker1.RunWorkerAsync();
            // DoWorkイベント発生
            Console.WriteLine("処理終了");

#if false
            List<String> list_data = new List<String>();
            if (comment_select.Text == "") ErrorAdd("「詳細設定」のコメントを選択してください。");
            if (title_text.Text == "") ErrorAdd("「詳細設定」のタイトル名を入力してください。");
            if (load_files_.Count == 0) ErrorAdd("転送元「SrcDir」ファイルがないですよ。");
            if (DstBox.Text == "") ErrorAdd("転送先「DstDir」フォルダがないですよ。");

            VoiceCellCheck_(1);     //キャラ名チェック
            VoiceCellCheck_(2);     //ラベル名チェック
            VoiceCellValueCheck_(3);//連番チェック
            VoiceCheck_();          //キャラ名禁止文字確認
            KeyCheck_();            //キャラデータ存在確認
            if (!ErrorCheck()) return;

            using (new KeepTime())
            {
                Libretto libretto = new Libretto();
                voice_data_.Clear();
                for (int i = 0; i < dataGridView1.Rows.Count; ++i)
                {
                    try
                    {
                        int serial = int.Parse(dataGridView1[3, i].Value.ToString());
                        int digit = dataGridView1[3, i].Value.ToString().Length;
                        voice_data_.Add(new global.VoiceData(
                            dataGridView1[1, i].Value.ToString(),
                            dataGridView1[2, i].Value.ToString(),
                            serial,
                            digit,
                            voice_top_text.Text,
                            voice_middle_text.Text,
                            voice_bottom_text.Text));
                    }
                    catch
                    {
                        ErrorAdd((i + 1).ToString() + "行目 : なんかエラーでてますよ");
                    }
                }
                if (!ErrorCheck()) return;

                //=============start!==============
                StartButton.Enabled = false;
                DetailSet();

                var excel_data = new global.ExcelData(100, 10, comment_select.Text, title_text.Text,
                                        check_voice_bold.Checked, color_text.ForeColor, color_text2.ForeColor,
                                        font_select.Text, version_text.Text, date_text.Text,
                                        convert_comment.Checked, check_add_star.Checked, check_visible.Checked);
                //progressBar1.Value
                libretto.Start(load_files_, DstBox.Text, voice_data_, excel_data, head_foot_pos_);

                //=================================

                StartButton.Enabled = true;
            }
#endif   
        }
        #endregion
        #region 転送先
        private void DstButton_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                DstBox.Text = folderBrowserDialog1.SelectedPath;
                HistoryList.Items.Add(DstBox.Text);
            }
        }

        private void DstBox_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.All;
            else
                e.Effect = DragDropEffects.None;
        }

        private void DstBox_DragDrop(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            if (Directory.Exists(s[0]))
            {
                DstBox.Text = s[0];
            }
            else
            {
                ErrorAdd(s[0] + " : フォルダを指定してください。");
            }
            if (!ErrorCheck()) return;
            HistoryList.Items.Add(s[0]);
            DstBox.Text = s[0];
        }

        private void DstBox_TextChanged(object sender, EventArgs e)
        {

        }
        #endregion
        #region 転送元
        private bool TextCheck_(string directory)
        {
            if (Directory.Exists(directory))
            {
                load_files_.Clear();
                src_list_box.Items.Clear();
                //フォルダの場合
                ArrayList files = new ArrayList();
                GetAllFiles(directory, "*.txt", ref files);
                HistoryList.Items.AddRange(files.ToArray());
                src_list_box.Items.AddRange(files.ToArray());
                if (files.Count == 0)
                {
                    ErrorAdd("テキストファイルが存在しないよ！");
                }
                else
                {
                    foreach (var file in files)
                    {
                        load_files_.Add(file.ToString());
                    }
                }
            }
            else
            {
                //ファイルの場合
                if (util.Util.FileCheck(directory, ".txt"))
                {
                    src_list_box.Items.Clear();
                    load_files_.Clear();
                    HistoryList.Items.Add(directory);
                    src_list_box.Items.Add(directory);
                    load_files_.Add(directory);
                }
                else
                {
                    ErrorAdd(directory + " : それテキストファイルじゃないよ！");
                }
            }
            if (!ErrorCheck()) return false;
            return true;
        }
        private void GetAllFiles(string folder, string searchPattern, ref ArrayList files)
        {
            string[] fs = System.IO.Directory.GetFiles(folder, searchPattern);
            files.AddRange(fs);

            string[] ds = System.IO.Directory.GetDirectories(folder);
            foreach (string d in ds)
                GetAllFiles(d, searchPattern, ref files);
        }
        private void SrcBox_TextChanged(object sender, EventArgs e)
        {

        }
        private void SrcButton_Click(object sender, EventArgs e)
        {
            
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                if (TextCheck_(folderBrowserDialog1.SelectedPath))
                {
                    AllClear();
                    SrcBox.Text = folderBrowserDialog1.SelectedPath;
                }
            }
        }
        private void SrcBox_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.All;
            else
                e.Effect = DragDropEffects.None;
        }
        private void SrcBox_DragDrop(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            if (TextCheck_(s[0]))
            {
                AllClear();
                SrcBox.Text = s[0];
            }
        }
        #endregion
        #region 履歴
        private void ClearButton_Click(object sender, EventArgs e)
        {
            HistoryList.Items.Clear();

            PlaySound();
            //StopSound();
        }
        #endregion
        #region ボイス設定 セル
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // 削除ボタン行かどうか確認 + 一番最後の *部分はクリック不可
            if (e.ColumnIndex == dataGridView1.Columns["Column1"].Index && e.RowIndex != -1)   //&& dataGridView1.Rows.Count - 1 != e.RowIndex
            {
                dataGridView1.Rows.RemoveAt(e.RowIndex);
            }
        }
        #endregion        
        #region XML設定ファイル書き込み・読み込み
        private MyXmlSettings xml = new MyXmlSettings();
        private void WriteXml(string path)
        {
            xml.SrcDir = SrcBox.Text;
            xml.DstDir = DstBox.Text;
            ArrayList[] data = new ArrayList[dataGridView1.Rows.Count];
            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                data[i] = new ArrayList();
                for (int j = 0; j < dataGridView1.ColumnCount; ++j)
                {
                    object s = dataGridView1[j, i].Value != null ? dataGridView1[j, i].Value: "";
                    data[i].Add(s.ToString());
                }
            }
            var cells = new string[data.Length][];
            for (int i = 0; i < data.Length; ++i)
            {
                List<string> s_list = new List<string>();
                for (int j = 0; j < dataGridView1.ColumnCount; ++j)
                {
                    s_list.Add(data[i][j].ToString());
                }
                cells[i] = s_list.ToArray();
            }
            xml.Cells = cells;
            xml.Comment = comment_select.Text;
            xml.Font = font_select.Text;
            xml.Title = title_text.Text;
            xml.LogoText = logo_text.Text;

            DetailSet();
            xml.LogoPos = head_foot_pos_[0].pos_hf * 3 + head_foot_pos_[0].pos_lcr;
            xml.PagePos = head_foot_pos_[1].pos_hf * 3 + head_foot_pos_[1].pos_lcr;
            xml.FilePos = head_foot_pos_[2].pos_hf * 3 + head_foot_pos_[2].pos_lcr;
            xml.KeyPos = head_foot_pos_[3].pos_hf * 3 + head_foot_pos_[3].pos_lcr;

            xml.VoiceBold = check_voice_bold.Checked;
            xml.VoiceOtherColor = ColorTranslator.ToOle(color_text.ForeColor);
            xml.VoiceColor = ColorTranslator.ToOle(color_text2.ForeColor);
            xml.Version = version_text.Text;

            xml.VoiceTop = voice_top_text.Text;
            xml.VoiceMiddle = voice_middle_text.Text;
            xml.VoiceBottom = voice_bottom_text.Text;

            xml.CommentConvert = convert_comment.Checked;
            xml.CheckStar = check_add_star.Checked;
            xml.Visible = check_visible.Checked;


            xml.CheckBox1 = checkBox1.Checked;
            xml.CheckBox2 = checkBox2.Checked;
            xml.CheckBox3 = checkBox3.Checked;
            xml.CheckBox4 = checkBox4.Checked;
            xml.CheckBox5 = checkBox5.Checked;

            xml.WriteXml(path);
        }
        private void ReadXml(string path)
        {
            if (!xml.ReadXml(path)) return;
            SrcBox.Text = xml.SrcDir;
            DstBox.Text = xml.DstDir;

            GridClear();
            foreach (var row in xml.Cells)
            {
                object[] col = row.ToArray();
                //dataGridView1.Rows.Add(row[0], row[1], row[2], row[3], row[4]);
                dataGridView1.Rows.Add(col);
            }

            comment_select.Text = xml.Comment;
            font_select.Text = xml.Font;
            title_text.Text = xml.Title;
            logo_text.Text = xml.LogoText;

            logo_select.SelectedIndex = xml.LogoPos;
            page_select.SelectedIndex = xml.PagePos;
            file_select.SelectedIndex = xml.FilePos;
            libretto_select.SelectedIndex = xml.KeyPos;

            voice_top_text.Text = xml.VoiceTop;
            voice_middle_text.Text = xml.VoiceMiddle;
            voice_bottom_text.Text = xml.VoiceBottom;

            check_voice_bold = CheckBoxRead(check_voice_bold, xml.VoiceBold);

            //ボイスキャラ以外の色
            color_text.ForeColor = ColorTranslator.FromOle(xml.VoiceOtherColor);
            if (color_text.ForeColor != Color.Black){
                other_voice_color = CheckBoxRead(other_voice_color, true);
                color_text.Text = GetStringRGB(color_text.ForeColor);
            }else{
                other_voice_color = CheckBoxRead(other_voice_color, false);
            }
            //ボイスキャラの色
            color_text2.ForeColor = ColorTranslator.FromOle(xml.VoiceColor);
            if (color_text2.ForeColor != Color.Black){
                voice_color = CheckBoxRead(voice_color, true);
                color_text2.Text = GetStringRGB(color_text2.ForeColor);
            }else{
                voice_color = CheckBoxRead(voice_color, false);
            }

            convert_comment = CheckBoxRead(convert_comment, xml.CommentConvert);
            check_add_star = CheckBoxRead(check_add_star, xml.CheckStar);
            check_visible = CheckBoxRead(check_visible, xml.Visible);
            version_text.Text = xml.Version;

            //日付デフォルト設定
            date_text.Text = util.Util.TodayUpdateString();
            // チェックボックス
            checkBox1.Checked = xml.CheckBox1;
            checkBox2.Checked = xml.CheckBox2;
            checkBox3.Checked = xml.CheckBox3;
            checkBox4.Checked = xml.CheckBox4;
            checkBox5.Checked = xml.CheckBox5;

            DetailSet();

            if (SrcBox.Text == "") return;
            TextCheckXml_(SrcBox.Text);
            HistoryList.Items.Add(DstBox.Text);
            src_list_box.Items.Clear();
            src_list_box.Items.AddRange(load_files_.ToArray());
        }
        private CheckBox CheckBoxRead(CheckBox src, bool check)
        {
            if (check)
            {
                src.Checked = true;
                src.CheckState = CheckState.Checked;
            }
            else
            {
                src.Checked = false;
                src.CheckState = CheckState.Unchecked;
            }
            return src;
        }

        private bool TextCheckXml_(string path)
        {
            
            load_files_.Clear();
            if (util.Util.FileCheck(path, ".txt"))
            {
                HistoryList.Items.Add(path);
                load_files_.Add(path);
            }
            else
            {
                var files = new ArrayList();
                GetAllFiles(path, "*.txt", ref files);
                if (files.Count == 0)
                {
                    ErrorAdd("テキストファイルが存在しないよ!");
                    return false;
                }
                HistoryList.Items.AddRange(files.ToArray());
                string[] data = new string[files.Count];
                Array.Copy(files.ToArray(), data, files.Count);
                load_files_.AddRange(data);
            }
            return true;
        }
        #endregion
        #region 詳細設定

        Hashtable detail_hash_ = new Hashtable();

        private System.Object KeySearch(Hashtable table, System.Object value)
        {
            foreach (var key in table.Keys)
            {
                if (table[key] == value)
                {
                    return key;
                }
            }
            return null;
        }
        private string[] hoge_ = new string[] { "0", "1", "2", "3", "4", "5", };
        private void Detailhoge()
        {
        }
        private void DetailInit()
        {
            //ポジション設定
            string[] detail_pos = new string[]
            {
                "0", "1", "2", "3", "4", "5",
            };


            logo_select.Items.AddRange(detail_pos);
            page_select.Items.AddRange(detail_pos);
            file_select.Items.AddRange(detail_pos);
            libretto_select.Items.AddRange(detail_pos);
            date_select.Items.AddRange(detail_pos);

            logo_select.SelectedIndex = 5;
            page_select.SelectedIndex = 3;
            file_select.SelectedIndex = 0;
            libretto_select.SelectedIndex = 2;
            date_select.SelectedIndex = 1;

            //フォント名設定
            string[] font_str = 
            {
                "ＭＳ Ｐゴシック",
                "ＭＳ Ｐ明朝",
                "ＭＳ ゴシック",
                "ＭＳ 明朝",
                "HG明朝E",
                "NSimSun",
                //"メイリオ",
            };
            font_select.Items.AddRange(font_str);
            font_select.SelectedIndex = 0;          //ＭＳ Ｐゴシックをデフォルト設定

            foreach (var font in font_str)
            {
                font_table_.Add(font);
            }

            //日付設定
            date_text.Text = util.Util.TodayUpdateString();
        }
        private void DetailSet()
        {
            int data = int.Parse(logo_select.SelectedIndex.ToString());
            var list = new[]
            {
                new { Pos = int.Parse(logo_select.SelectedIndex.ToString()), Text = logo_text.Text },
                new { Pos = int.Parse(page_select.SelectedIndex.ToString()), Text = "" },
                new { Pos = int.Parse(file_select.SelectedIndex.ToString()), Text = "" },
                new { Pos = int.Parse(libretto_select.SelectedIndex.ToString()), Text = "" },
                new { Pos = int.Parse(date_select.SelectedIndex.ToString()), Text = "" },
            };
            for (int i = 0; i < head_foot_pos_.Length; i++)
			{
                head_foot_pos_[i] = SetHeadFoot(list[i].Pos, list[i].Text);
			}
        }
        private global.HeadFootPos SetHeadFoot(int posY, int posX, string text)
        {
            return new global.HeadFootPos(posY, posX, text);
        }
        private global.HeadFootPos SetHeadFoot(int pos, string text)
        {
            return new global.HeadFootPos(pos / 3, pos % 3, text);
        }
        private HashSet<string> font_table_ = new HashSet<string>();
#region 画像表示テスト
#if false
        private void button1_Click(object sender, EventArgs e)
        {
            if (File.Exists(picture_text.Text))
            {
                MyExcel excel = new MyExcel();
                excel.Visible();
                excel.CreateBook(1);

                System.Drawing.Bitmap bmp = new System.Drawing.Bitmap(picture_text.Text);
                float width = bmp.Width;
                float height = bmp.Height;
                excel.PictureExpantion(picture_text.Text, 0, 0, width, height);
                excel.End();
            }
            else
            {
                if (picture_text.Text == "")
                {
                    ErrorAdd("ファイルをドラッグ&ドロップしてください。");
                }
                else
                {
                    ErrorAdd(picture_text.Text + " : ファイルが存在しません！");
                }
                ErrorCheck();
            }
        }
        private void picture_text_DragDrop(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            if (PictureCheck_(s[0]))
            {
                picture_text.Text = s[0];
            }

        }
        private void picture_text_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.All;
            else
                e.Effect = DragDropEffects.None;
        }
        private bool PictureCheck_(string path)
        {
            string s = System.IO.Path.GetExtension(path);
            string[] extentions = { ".bmp", ".gif", ".jpeg", ".jpg", ".png" };
            foreach (var et in extentions)
            {
                if (et == s)
                {
                    return true;
                }
            }
            return false;
        }
#endif
#endregion
        private void color_button_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                //バックカラーを変更するのは、文字色を見やすくするためである。それとReadOnly = true　の状態だとTextBox.ForeColorが反映されないので
                //http://bokuibi.blogspot.jp/2009/06/netreadonlytruetextboxforecolor.html
                byte[] color = new byte[3];
                color[0] = colorDialog1.Color.R;
                color[1] = colorDialog1.Color.G;
                color[2] = colorDialog1.Color.B;

                List<bool> back_reverse = new List<bool>();
                foreach (var col in color)
                {
                    if (col >= 128)
                    {
                        back_reverse.Add(true);
                    }
                }
                if (back_reverse.Count >= 2)
                {
                    color_text.BackColor = Color.Black;
                }
                else
                {
                    color_text.BackColor = Color.White;
                }
                color_text.ForeColor = colorDialog1.Color;
                color_text.Text = GetStringRGB(color[0], color[1], color[2]);
            }
        }
        private void color_button2_Click_1(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                //バックカラーを変更するのは、文字色を見やすくするためである。それとReadOnly = true　の状態だとTextBox.ForeColorが反映されないので
                //http://bokuibi.blogspot.jp/2009/06/netreadonlytruetextboxforecolor.html
                byte[] color = new byte[3];
                color[0] = colorDialog1.Color.R;
                color[1] = colorDialog1.Color.G;
                color[2] = colorDialog1.Color.B;

                List<bool> back_reverse = new List<bool>();
                foreach (var col in color)
                {
                    if (col >= 128)
                    {
                        back_reverse.Add(true);
                    }
                }
                if (back_reverse.Count >= 2)
                {
                    color_text2.BackColor = Color.Black;
                }
                else
                {
                    color_text2.BackColor = Color.White;
                }
                color_text2.ForeColor = colorDialog1.Color;
                color_text2.Text = GetStringRGB(color[0], color[1], color[2]);
            }
        }
        private string GetStringRGB(Color color)
        {
            return "R:" + color.R.ToString() + " G:" + color.G.ToString() + " B:" + color.B.ToString();
        }
        private string GetStringRGB(byte r, byte g, byte b)
        {
            return "R:" + r.ToString() + " G:" + g.ToString() + " B:" + b.ToString();
        }
        private void check_no_voice_CheckedChanged(object sender, EventArgs e)
        {
            if (other_voice_color.Checked)
            {
                color_button.Visible = true;
                color_text.Visible = true;
                color_text.Text = GetStringRGB(color_text.ForeColor);
            }
            else
            {
                color_button.Visible = false;
                color_text.Visible = false;
                color_text.ForeColor = Color.Black;
                color_text.BackColor = Color.White;
                color_text.Text = "";
            }
        }
        private void voice_color_CheckedChanged_1(object sender, EventArgs e)
        {
            if (voice_color.Checked)
            {
                color_button2.Visible = true;
                color_text2.Visible = true;
                color_text2.Text = GetStringRGB(color_text2.ForeColor);
            }
            else
            {
                color_button2.Visible = false;
                color_text2.Visible = false;
                color_text2.ForeColor = Color.Black;
                color_text2.BackColor = Color.White;
                color_text2.Text = "";
            }
        }
        private void font_select_SelectedIndexChanged(object sender, EventArgs e)
        {
            font_select.Font = new Font(font_select.Text, 9);
        }
        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            switch (e.ColumnIndex)
            {
                case 1:
                    VoiceCellCheck_(e.ColumnIndex);
                    break;
                case 2:
                    VoiceCellCheck_(e.ColumnIndex);
                    break;
                case 3:
                    VoiceCellValueCheck_(e.ColumnIndex);
                    break;
            }
            ErrorCheck();
        }
        #endregion
        #region 音声スタイル
        private void ChangeVoiceStyle()
        {
            var s = voice_top_text.Text + voice_label_text.Text + voice_middle_text.Text + voice_number_text.Text + voice_bottom_text.Text;

            voice_text.Text = s;
        }
        private void ChangeVoiceStyle(string label, string number)
        {
            voice_label_text.Text = label;
            voice_number_text.Text = number;
            ChangeVoiceStyle();
        }
        private void voice_top_text_TextChanged(object sender, EventArgs e)
        {
            ChangeVoiceStyle();
        }
        private void voice_label_text_TextChanged(object sender, EventArgs e)
        {
            ChangeVoiceStyle();
        }
        private void voice_middle_text_TextChanged(object sender, EventArgs e)
        {
            ChangeVoiceStyle();
        }
        private void voice_number_text_TextChanged(object sender, EventArgs e)
        {
            ChangeVoiceStyle();
        }
        private void voice_bottom_text_TextChanged(object sender, EventArgs e)
        {
            ChangeVoiceStyle();
        }
        private void CheckCell(int x, int y)
        {
            //if (y == -1 || y >= dataGridView1.RowCount - 1) return;
            if (y == -1) return;
            if (dataGridView1[2, y].Value == null || dataGridView1[3, y].Value == null) return;

            ChangeVoiceStyle(dataGridView1[2, y].Value.ToString(), dataGridView1[3, y].Value.ToString());
        }
        private void dataGridView1_CellMouseEnter(object sender, DataGridViewCellEventArgs e)
        {
            var x = e.ColumnIndex;
            var y = e.RowIndex;
            CheckCell(x, y);
        }
        private void dataGridView1_CellMouseMove(object sender, DataGridViewCellMouseEventArgs e)
        {
            var x = e.ColumnIndex;
            var y = e.RowIndex;
            CheckCell(x, y);
        }
        private void dataGridView1_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            var x = e.ColumnIndex;
            var y = e.RowIndex;
            CheckCell(x, y);
        }
        
        // ボイス検索
        private void voice_search_Click(object sender, EventArgs e)
        {
            SetKey_(comment_select.Text);
            ErrorCheck();
        }
        #endregion
        #region メニューストリップ
        private void file_read_Click(object sender, EventArgs e)
        {
            //初期ディレクトリ設定
            openFileDialog1.InitialDirectory = System.Environment.CurrentDirectory;

            //開く
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                ReadXml(openFileDialog1.FileName);
            }
        }
        private void file_save_Click_1(object sender, EventArgs e)
        {
            saveFileDialog1.InitialDirectory = System.Environment.CurrentDirectory;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //保存
                WriteXml(saveFileDialog1.FileName);
            }
        }
        #endregion

        private void convert_comment_CheckedChanged(object sender, EventArgs e)
        {
            //チェック判定 特になし
        }

        private void comment_select_SelectedIndexChanged(object sender, EventArgs e)
        {
            convert_comment.Text = comment_select.Text + "(先頭コメント) → ◇(変換)";
        }

        #region スレッドテスト
#if false
        //進捗処理用デリゲート
        private delegate void SetProgressDelegate();
        //終了時用デリゲート
        private delegate void ThreadCompletedDelegate();
        //キャンセル用デリゲート
        private delegate void ThreadCanceldDelegate();
        //スレッド用変数
        private volatile bool canceled = false;
        //スレッド
        private System.Threading.Thread workerThread;
        private void ThreadStart()
        {
            progressBar1.Value = 0;
            progressBar1.Minimum = 0;
            progressBar1.Maximum = voice_data_.Count;

            workerThread = new System.Threading.Thread(new System.Threading.ThreadStart(CountUp));
            workerThread.IsBackground = true;
            workerThread.Start();
        }
        private void CountUp()
        {
            var progressDlg = new SetProgressDelegate(SetProgressValue);
            var completed = new ThreadCompletedDelegate(ThreadCompleted);
            var canceld = new ThreadCanceldDelegate(ThreadCanceled);

            while(progressBar1.Value != voice_data_.Count)
            {
                //this.Invoke(progressDlg, new object[] { 3 });
                this.Invoke(progressDlg);
            }
            //完了したときにコントロールの値を変更する
            this.Invoke(completed);
        }
        private void SetProgressValue()
        {
            progressBar1.Value = progress_;
        }
        private void ThreadCompleted()
        {
            MessageBox.Show("終了");
        }
        private void ThreadCanceled()
        {
            MessageBox.Show("キャンセル");
        }
#endif
        #endregion
        private void program_help_Click(object sender, EventArgs e)
        {
            //Help.chm
            try
            {
                Process p = Process.Start(Directory.GetCurrentDirectory() + "/Help.chm");
            }
            catch
            {
                ErrorAdd("Not Found!");
                ErrorCheck();
            }
        }

        private void debugButton_Click(object sender, EventArgs e)
        {
            foreach(var file in load_files_)
                dataGridView2.Rows.Add(file);

            List<String> list_data = new List<String>();
            if (load_files_.Count == 0) ErrorAdd("転送元「SrcDir」ファイルがないですよ。");
            if (DstBox.Text == "") ErrorAdd("転送先「DstDir」フォルダがないですよ。");
            VoiceCellCheck_(1);     //キャラ名チェック
            VoiceCellCheck_(2);     //ラベル名チェック
            VoiceCellValueCheck_(3);//連番チェック
            FileNameCheck();
            VoiceCheck_();          //キャラ名禁止文字確認
            KeyCheck_();            //キャラデータ存在確認
            if (!ErrorCheck()) return;
            voice_data_.Clear();
            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                try
                {
                    int serial = int.Parse(dataGridView1[3, i].Value.ToString());
                    int digit = dataGridView1[3, i].Value.ToString().Length;
                    voice_data_.Add(new global.VoiceData(
                        dataGridView1[1, i].Value.ToString(),
                        dataGridView1[2, i].Value.ToString(),
                        serial,
                        digit,
                        voice_top_text.Text,
                        voice_middle_text.Text,
                        voice_bottom_text.Text, false));
                }
                catch
                {
                    ErrorAdd((i + 1).ToString() + "行目 : なんかエラーでてますよ");
                }
            }

            // エラーチェック
            if (!ErrorCheck()) return;
            libretto.DebugLog(load_files_, DstBox.Text, voice_data_);

            //Help.chm
            try
            {
                Process p = Process.Start(DstBox.Text + "/debug.log");
            }
            catch
            {
                ErrorAdd("log失敗");
                ErrorCheck();
            }
        }

        private void TransparentToolStrip_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        #region 出力設定
        // チェックされたとき
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            out_flg_.libretto_omission = checkBox1.Checked;
        }
        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            out_flg_.voice_excel= checkBox2.Checked;

        }
        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            out_flg_.libretto= checkBox3.Checked;
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            out_flg_.lib_only = checkBox4.Checked;
        }
        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            out_flg_.voice_lib_only = checkBox5.Checked;
        }
        private void checkBox2_MouseEnter(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            var s = "ぬき本を作成するかしないか on = する off = しない";
            listBox1.Items.Add(s);
        }

        private void checkBox3_MouseEnter(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            var s = "キャラ台本を作成するかしないか on = する off = しない";
            listBox1.Items.Add(s);
        }

        private void checkBox4_MouseEnter(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            var s = "全体のボイス番号無しの台本を作成するかしないか on = する off = しない";
            listBox1.Items.Add(s);
        }
        private void checkBox5_MouseEnter(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            var s = "全体のボイス番号有りの台本を作成するかしないか on = する off = しない";
            listBox1.Items.Add(s);
        }
        #endregion 
        #region MUSICテスト
        private System.Media.SoundPlayer player = null;
        private void PlaySound()
        {

            string dir = System.IO.Directory.GetCurrentDirectory();
            string sound_dir = dir + "/sound/";
            if (!System.IO.Directory.Exists(sound_dir))
                return;

            string[] files = System.IO.Directory.GetFiles(sound_dir, "*.wav");
            if (files.Length == 0)
                return;

            //再生されているときは止める
            if (player != null)
                StopSound();

            //読み込む
            player = new System.Media.SoundPlayer(files[0]);
            //非同期再生する
            player.Play();
        }
        //再生されている音を止める
        private void StopSound()
        {
            if (player != null)
            {
                player.Stop();
                player.Dispose();
                player = null;
            }
        }

        #endregion

        #region デバッグ
        private void textBox1_DragDrop(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            List<TreeNode> list = new List<TreeNode>();
            foreach (var dir in s)
            {
                var tree = new TreeNode(dir);
                if (Directory.Exists(dir))
                {
                    var directory = new DirectoryInfo(dir);
                    DirectoryHierarchy(tree, directory);
                }
                tree.Text = Path.GetFileName(dir);
                list.Add(tree);
            }
            treeView1.Nodes.Clear();
            treeView1.Nodes.AddRange(list.ToArray());
        }

        private void textBox1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.All;
            else
                e.Effect = DragDropEffects.None;
        }
        private void DirectoryHierarchy(TreeNode parent, DirectoryInfo dir)
        {
            // ディレクトリ
            foreach (var dir2 in dir.GetDirectories())
            {
                TreeNode child = new TreeNode(dir2.Name, 1, 2);
                parent.Nodes.Add(child);
                DirectoryHierarchy(child, dir2);
            }
            // ファイル
            foreach (var file in dir.GetFiles())
            {
                System.Text.Encoding enc = System.Text.Encoding.Default;
                // テキストファイルのみ追加
                if (file.Extension == ".txt")
                {
                    var fs = new System.IO.FileStream(file.DirectoryName + "/" + file.Name, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                    byte[] bs = new byte[fs.Length];
                    fs.Read(bs, 0, bs.Length);
                    fs.Close();
                    enc = util.Util.GetCode(bs);
                }

                TreeNode child = new TreeNode(file.Name + " : " + enc.EncodingName, 3, 3);
                parent.Nodes.Add(child);
            }
        }
        #endregion

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }


    }
}

