using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;

using System.Windows.Forms;
using System.ComponentModel;

using System.Threading.Tasks;
using System.Threading;

namespace LibrettoCreateTool
{
    /// <summary>
    /// 外部で使用する唯一の関数
    /// </summary>
    public partial class Libretto
    {

        // 大村さんのScenarioCollectionで作成されたテキストを通常のテキストに戻す。
        // 実はこれをやるより、もともとのデータをバックアップで持っていたほうが絶対にエラーが出ない。
        public void ReConvertScenarioCollection(List<string> src_files)
        {
            string[][] file_lines = src_files.Select(file => File.ReadAllLines(file, Encoding.Default)).ToArray();
            string[] file_names = src_files.ToArray();

            List<string> result = new List<string>();
            for(int i = 0; i < file_lines.Length; ++i)
            {
                for(int j = 0; j < file_lines[i].Length; ++j)
                {
                    if (file_lines[i][j].StartsWith("#!") || file_lines[i][j].StartsWith("#?"))
                    {
                        file_lines[i][j] = file_lines[i][j].Substring(13);
                    }
                }
                var sb = new StringBuilder();
                file_lines[i].Select(line => sb.AppendLine(line)).ToArray();
                //result.Add(sb.ToString());
                using (var sw = new StreamWriter(file_names[i], false, Encoding.Default))
                {
                    sw.Write(sb.ToString());
                }
            }
        }

        public void Start(List<string> src_files, string dst_dir, List<global.VoiceData> voice_data, global.ExcelData excel_data, global.HeadFootPos[] detail_data, global.OutputFlags flgs, BackgroundWorker worker)
        {
            
            Initialize_();
            // 関数の奥で使うため、毎回excel_data.conver_commentのためのbool引数を追加するのが面倒なので、ここで取得してます。
            comment_convert_ = excel_data.convert_comment;
            // omissionも内部で使うので取得しときます。
            omission_ = flgs.libretto_omission;
            CreateOverlapLabel_(voice_data);
            FileAdd_(src_files.ToArray(), excel_data.convert_comment, excel_data.comment);

            var directorys = new string[7]
            {
                dst_dir + "/libretto_xls/",
                dst_dir + "/libretto_pdf/",
                dst_dir + "/libretto_key/",

                dst_dir + "/libretto_txt/",
                dst_dir + "/libretto_BAK/",

                dst_dir + "/通し台本/" + "libretto/",
                dst_dir + "/通し台本/" + "key/",
            };
            if (!CreateLibrettoDirectorys_(directorys)) return;
            foreach (var file in src_files)
            {
                CreateBackUpText_(file, directorys[4]);
            }

            CreateConvertText_(voice_data, directorys[3]);

            MyExcel excel = new MyExcel();
            if (excel_data.excel_visible) excel.Visible();
#if true
            if (flgs.voice_excel)
            {
                if (!CreateVoiceExcel_(excel, voice_data, excel_data, directorys[2])) return;
            }
            if (flgs.libretto)
            {
                if (!CreateLibrettoExcelAndPDF_(excel, voice_data, excel_data, detail_data, directorys, worker)) return;
            }
            if (flgs.lib_only)
            {
                if (!test2_(excel, voice_data, excel_data, detail_data, directorys, worker)) return;
            }
#endif
            // 通し台本は固定で作成
            CreateLabelHash(voice_data);

            if (flgs.voice_lib_only)
            {
                CreateAllPDF(excel, directorys[5], excel_data, detail_data);
            }
            CreateAllPDFCharacter(excel, directorys[5], excel_data, detail_data);

            CreateVoiceExcel(excel, excel_data, directorys[6]);

            excel.End();
        }
        public void Test()
        {
#if false
            MyExcel excel = new MyExcel();
            excel.Visible();
            excel.CreateBook(1);
            int all = 0;
            Parallel.For(0, 51, i => {
                excel.CreateNewSheet(i.ToString());
                object[,] cell = new object[10 + 1, 7];
                cell[0, 0] = "収録数";
                cell[0, 1] = "シーン名";
                cell[0, 2] = "台本項";
                cell[0, 3] = "音声番号";
                cell[0, 4] = "備考";
                cell[0, 5] = "台詞(収録後)";
                cell[0, 6] = "台詞(収録前)";
                excel.CreateCell(0, 0, cell);
                excel.SetRange();
                excel.SetFontColor(MyExcel.EXCEL_MATRIX.X, 0, Color.White);
                excel.SetInteriorColor(MyExcel.EXCEL_MATRIX.X, 0, Color.Green);
                //excel.SetWidthAutoFit();

                excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.Y, 0, 7);
                excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.Y, 1, 25);
                excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.Y, 2, 7);
                excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.Y, 3, 17);
                excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.Y, 4, 7);
                excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.Y, 5, 64);
                excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.Y, 6, 64);
            });
            excel.Close();

            excel.End();
#endif
        }
    }
    /// <summary>
    /// 拡張 05/21
    /// 全てのボイスを作成する台本1つ
    /// </summary>
    public partial class Libretto
    {
        void CreateVoiceExcel(MyExcel excel, global.ExcelData excel_data, string directory)
        {
            List<string> character_list = new List<string>();
            foreach (global.VoiceData voice_data in label_hash_.Values)
            {
                if (voice_data.serial_lib)
                    character_list.Add((string)overlap_label_[voice_data.key]);
            }

            foreach (string character in character_list)
            {
                var lk = new List<LibrettoKey>();
                excel.CreateBook(1);
                excel.SetZoom(80);
                int page_max = OmissionFileMax(lib_);
                int page_count = 2;
                int all_count = 0;

                foreach (var file in lib_.files)
                {
                    string file_name = Path.GetFileNameWithoutExtension(file.file_name);
                    int str_cat_index = -1;
                    foreach (var page in file.pages)
                    {
                        ++page_count;
                        foreach (var line in page.lines)
                        {

                            if (character == (string)overlap_label_[line.key])
                            {
                                if (line.over_lap) continue;
                                global.VoiceData voice_data = (global.VoiceData)label_hash_[character];
                                var s_serial = voice_data.top + CreateVoiceFormat(line.v_count, voice_data) + voice_data.bottom;
                                lk.Add(new LibrettoKey(++all_count, file_name, page_count, s_serial, character, line.serif));
                                if (line.serif.IndexOf("」") != -1)
                                    str_cat_index = -1;
                                else
                                    str_cat_index = lk.Count - 1;
                            }
                            else if (str_cat_index != -1)
                            {
                                lk[str_cat_index].serif += line.serif;
                                if (line.serif.IndexOf("」") != -1) str_cat_index = -1;
                            }
                        }
                    }
                    ++page_count;
                }

                object[,] cell = new object[lk.Count + 1, 7];
                cell[0, 0] = "収録数";
                cell[0, 1] = "シーン名";
                cell[0, 2] = "台本項";
                cell[0, 3] = "音声番号";
                cell[0, 4] = "備考";
                cell[0, 5] = "台詞(収録後)";
                cell[0, 6] = "台詞(収録前)";
                for (int i = 0; i < lk.Count; ++i)
                {
                    cell[i + 1, 0] = lk[i].all_count;
                    cell[i + 1, 1] = lk[i].file_name;
                    cell[i + 1, 2] = lk[i].detail_count;
                    cell[i + 1, 3] = lk[i].s_serial;
                    //cell[i + 1, 4] = " ";
                    cell[i + 1, 5] = lk[i].serif;
                    cell[i + 1, 6] = cell[i + 1, 5];
                }
                excel.CreateCell(0, 0, cell);
                excel.SetRange();
                excel.SetFontColor(MyExcel.EXCEL_MATRIX.X, 0, Color.White);
                excel.SetInteriorColor(MyExcel.EXCEL_MATRIX.X, 0, Color.Green);
                excel.SetFontName(excel_data.font_name);    //デフォルトフォントなら処理しない方が断然早い
                //excel.SetWidthAutoFit();

                excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.Y, 0, 7);
                excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.Y, 1, 25);
                excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.Y, 2, 7);
                excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.Y, 3, 17);
                excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.Y, 4, 7);
                excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.Y, 5, 64);
                excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.Y, 6, 64);
                excel.SetBorderColor(Color.Navy);
                excel.CellWrapMode();
                //PDF作成
                excel.PrintOrientation();
                excel.CreatePDF(directory + character + "_抜き台本.xlsx", 0, 5);

                //ブック毎にセーブ xlsx作成
                excel.SaveXlsx(directory + character + "_抜き台本.xlsx");
                excel.Close();
            }
        }

        // 全ボイス台本用カバー
        void AllCover(MyExcel excel, global.ExcelData excel_data, global.HeadFootPos[] detail_data, ref int sheet_count)
        {
            int page_max = OmissionFileMax(lib_);
            
            //セル設定
            excel.CreateNewSheet((++sheet_count).ToString());
            excel.CreateCell(0, 0, 3, 1);

            excel.SetCell(0, 0, excel_data.title);
            excel.SetCell(1, 0, "バージョン : " + excel_data.version);
            excel.SetCell(2, 0, "日付 : " + excel_data.today_date);

            excel.SetRange();
            excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.X, 0, 50);
            excel.CellAlignment(MyExcel.CELL_H_DIR.CENTER, MyExcel.EXCEL_MATRIX.Y, 0);
            excel.SetFontSize(26);
            excel.SetWidthAutoFit();
            excel.SetHeightAutoFit();

            //印刷設定
            excel.PrintMiddle();
            excel.PrintOrientation();
            excel.PrintHeaderFooter(detail_data[1].pos_hf, detail_data[1].pos_lcr, sheet_count.ToString() + "/" + page_max.ToString());
            excel.PrintSize(MyExcel.PRINT_SIZE.A4);
        }
        // 全ボイスの台本作成
        void CreateAllPDF(MyExcel excel, string path, global.ExcelData excel_data, global.HeadFootPos[] detail_data)
        {
            excel.CreateBook(1);
            int page_max = OmissionFileMax(lib_);
            int sheet_count = 0;

            // 表紙
            AllCover(excel, excel_data, detail_data, ref sheet_count);

            // 中身
            foreach (var files in lib_.files)
            {
                excel.CreateNewSheet((++sheet_count).ToString());
                excel.CreateCell(0, 0, 2, 1);
                excel.SetCell(0, 0, "シーン : " + Path.GetFileNameWithoutExtension(files.file_name));
                excel.SetCell(1, 0, "台本");
                excel.SetRange();
                excel.SetFontSize(22);
                //印刷設定
                excel.CellAlignment(MyExcel.CELL_H_DIR.CENTER, MyExcel.EXCEL_MATRIX.Y, 0);
                excel.PrintMiddle();
                excel.PrintOrientation();
                excel.SetWidthAutoFit();
                excel.PrintHeaderFooter(detail_data[1].pos_hf, detail_data[1].pos_lcr, sheet_count.ToString() + "/" + page_max.ToString());
                excel.PrintSize(MyExcel.PRINT_SIZE.A4);

                foreach (var page in files.pages)
                {
                    excel.CreateNewSheet((++sheet_count).ToString());
                    excel.CreateCell(0, 0, PAGE_LINE_MAX + 1, 3);
                    excel.SetCell(0, 0, "音声番号");
                    excel.SetCell(0, 1, "キャラ名");
                    excel.SetCell(0, 2, "セリフ");
                    int i = 0;
                    foreach (var line in page.lines)
                    {
                        //excel 書き込み
                        if (line.v_count != -1)
                        {
                            if(overlap_label_.ContainsKey(line.key))
                            {
                                global.VoiceData v = (global.VoiceData)label_hash_[overlap_label_[line.key]];
                                excel.SetCell(i + 1, 0, CreateVoiceFormat(line.v_count, v));
                            }
                        }

                        if (overlap_label_.ContainsKey(line.key))
                            excel.SetCell(i + 1, 1, overlap_label_[line.key]);
                        else
                            excel.SetCell(i + 1, 1, line.key);
                            
                        excel.SetCell(i + 1, 2, line.serif);
                        ++i;
                    }

                    // 設定
                    excel.ChangeXY();
                    excel.SetRange();
                    //縦書き設定
                    excel.SellxlVertical();
                    //ボーダー設定
                    Color border_color = Color.Green;
                    excel.BorderRange(MyExcel.EXCEL_MATRIX.Y, excel.EndX() - 1, border_color);
                    excel.BorderRange(MyExcel.BORDER_RANGE.EDGE_BOTTOM, border_color);
                    excel.BorderRange(MyExcel.BORDER_RANGE.EDGE_LEFT, border_color);
                    excel.BorderRange(MyExcel.BORDER_RANGE.EDGE_TOP, border_color);
                    excel.BorderRange(MyExcel.BORDER_RANGE.INSIDE_HORIZONTAL, border_color);
                    excel.SellxOrientation(-90, MyExcel.EXCEL_MATRIX.X, 0, excel.StartX(), excel.EndX() - 2);  //ボイス行のみ横文字、横表示
                    excel.CellAlignment(MyExcel.CELL_H_DIR.CENTER, MyExcel.EXCEL_MATRIX.X, 0);        //ボイス行のみ、文字詰めセンター調整
                    excel.SetFontName(excel_data.font_name);    //デフォルトフォントなら処理しない方が断然早い
                    //各種セル手動調整
                    excel.SetFontSize(10);
                    excel.CellAlignment(MyExcel.CELL_V_DIR.TOP);
                    excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.X, 0, 3);
                    excel.CellSizeHeight(MyExcel.EXCEL_MATRIX.X, 0, 70);
                    excel.CellSizeHeight(MyExcel.EXCEL_MATRIX.X, 1, 80);
                    excel.CellSizeHeight(MyExcel.EXCEL_MATRIX.X, 2, 385);
                    //印刷設定
                    excel.PrintOrientation();
                    excel.PrintArea();
                    excel.PrintSize(MyExcel.PRINT_SIZE.A4);
                    string[] page_heder_footer = { 
                        detail_data[0].print,                                       //ロゴ
                        (sheet_count).ToString() + "/" + (page_max),                 //ページ数 最初の1ページは削除するため
                        Path.GetFileNameWithoutExtension(files.file_name),       //ファイル名
                        "台本",                                                   //台本名
                        DateTime.Today.ToShortDateString()                          //日付
                    };

                    for (int k = 0; k < page_heder_footer.Length; k++)
                    {
                        excel.PrintHeaderFooter(detail_data[k].pos_hf, detail_data[k].pos_lcr, page_heder_footer[k]);
                    }
                }
            }

            excel.SheetDelete(1);
            excel.SaveXls(path + "台本テスト.xls");
            excel.CreatePDF(path + "台本テスト.xls");
            excel.Close();
        }

        void CreateAllPDFCharacter(MyExcel excel, string path, global.ExcelData excel_data, global.HeadFootPos[] detail_data)
        {
            List<string> character_list = new List<string>();
            foreach (global.VoiceData voice_data in label_hash_.Values)
            {
                if (voice_data.serial_lib)
                    character_list.Add((string)overlap_label_[voice_data.key]);
            }

            foreach (string character in character_list)
            {
                excel.CreateBook(1);
                int page_max = OmissionFileMax(lib_);
                int sheet_count = 0;

                // 表紙
                AllCover(excel, excel_data, detail_data, ref sheet_count);    

                // 中身
                foreach (var files in lib_.files)
                {
                    excel.CreateNewSheet((++sheet_count).ToString());
                    excel.CreateCell(0, 0, 2, 1);
                    excel.SetCell(0, 0, "シーン : " + Path.GetFileNameWithoutExtension(files.file_name));
                    excel.SetCell(1, 0, "台本");
                    excel.SetRange();
                    excel.SetFontSize(22);
                    //印刷設定
                    excel.CellAlignment(MyExcel.CELL_H_DIR.CENTER, MyExcel.EXCEL_MATRIX.Y, 0);
                    excel.PrintMiddle();
                    excel.PrintOrientation();
                    excel.SetWidthAutoFit();
                    excel.PrintHeaderFooter(detail_data[1].pos_hf, detail_data[1].pos_lcr, sheet_count.ToString() + "/" + page_max.ToString());
                    excel.PrintSize(MyExcel.PRINT_SIZE.A4);

                    bool str_cat = false;
                    bool over_lap_cat = false;

                    foreach (var page in files.pages)
                    {
                        excel.CreateNewSheet((++sheet_count).ToString());
                        excel.CreateCell(0, 0, PAGE_LINE_MAX + 1, 3);
                        excel.SetCell(0, 0, "音声番号");
                        excel.SetCell(0, 1, "キャラ名");
                        excel.SetCell(0, 2, "セリフ");
                        List<int> voice_index = new List<int>();
                        List<int> over_lap_index = new List<int>();

                        int i = 0;
                        foreach (var line in page.lines)
                        {
                            //excel 書き込み
                            if (character == (string)overlap_label_[line.key])
                            {
                                excel.SetCell(i + 1, 0, CreateVoiceFormat(line.v_count, (global.VoiceData)label_hash_[character]));
                                if(excel_data.check_star)
                                    excel.SetCell(i + 1, 1, "☆" + character);
                                else
                                    excel.SetCell(i + 1, 1, character);
                                excel.SetCell(i + 1, 2, line.serif);

                                if (line.over_lap)
                                {
                                    if (line.serif.IndexOf("」") != -1)
                                        over_lap_cat = false;
                                    else
                                        over_lap_cat = true;
                                    over_lap_index.Add(i + 1);
                                }
                                else 
                                {
                                    if (line.serif.IndexOf("」") != -1)
                                        str_cat = false;
                                    else
                                        str_cat = true;
                                    voice_index.Add(i + 1);
                                }
                            }
                            else
                            {
                                excel.SetCell(i + 1, 0, "");
                                excel.SetCell(i + 1, 1, line.key);
                                excel.SetCell(i + 1, 2, line.serif);
                            }
                            if (str_cat)
                            {
                                voice_index.Add(i + 1);
                                if (line.serif.IndexOf("」") != -1)
                                    str_cat = false;
                            }
                            if (over_lap_cat)
                            {
                                over_lap_index.Add(i + 1);
                                if (line.serif.IndexOf("」") != -1)
                                    over_lap_cat = false;
                            }
                            ++i;
                        }

                        // 設定
                        excel.ChangeXY();
                        excel.SetRange();

                        // 全体のフォント色
                        SetFontColor_(excel, excel_data.other_voice_color);
                        // カブリボイス色 灰色固定
                        SetFontColor_(excel, over_lap_index, Color.Gray);
                        // ボイス色 ※太字なら太字追加
                        if (excel_data.bold)
                            SetFontBold_(excel, voice_index, excel_data.voice_color);
                        else
                            SetFontColor_(excel, voice_index, excel_data.voice_color);

                        excel.SellxlVertical();                     //縦書き設定
                        excel.SetFontName(excel_data.font_name);    //フォント設定



                        //ボーダー設定
                        Color border_color = Color.Green;
                        excel.BorderRange(MyExcel.EXCEL_MATRIX.Y, excel.EndX() - 1, border_color);
                        excel.BorderRange(MyExcel.BORDER_RANGE.EDGE_BOTTOM, border_color);
                        excel.BorderRange(MyExcel.BORDER_RANGE.EDGE_LEFT, border_color);
                        excel.BorderRange(MyExcel.BORDER_RANGE.EDGE_TOP, border_color);
                        excel.BorderRange(MyExcel.BORDER_RANGE.INSIDE_HORIZONTAL, border_color);
                        excel.SellxOrientation(-90, MyExcel.EXCEL_MATRIX.X, 0, excel.StartX(), excel.EndX() - 2);  //ボイス行のみ横文字、横表示
                        excel.CellAlignment(MyExcel.CELL_H_DIR.CENTER, MyExcel.EXCEL_MATRIX.X, 0);        //ボイス行のみ、文字詰めセンター調整

                        //各種セル手動調整
                        excel.SetFontSize(10);
                        excel.CellAlignment(MyExcel.CELL_V_DIR.TOP);
                        excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.X, 0, 3);
                        excel.CellSizeHeight(MyExcel.EXCEL_MATRIX.X, 0, 70);
                        excel.CellSizeHeight(MyExcel.EXCEL_MATRIX.X, 1, 80);
                        excel.CellSizeHeight(MyExcel.EXCEL_MATRIX.X, 2, 385);
                        //印刷設定
                        excel.PrintOrientation();
                        excel.PrintArea();
                        excel.PrintSize(MyExcel.PRINT_SIZE.A4);
                        string[] page_heder_footer = {
                            detail_data[0].print,                                       //ロゴ
                            (sheet_count).ToString() + "/" + (page_max),                //ページ数 最初の1ページは削除するため
                            Path.GetFileNameWithoutExtension(files.file_name),       //ファイル名
                            character + "台本",                                         //台本名
                            DateTime.Today.ToShortDateString()                          //日付
                        };

                        for (int k = 0; k < page_heder_footer.Length; k++)
                        {
                            excel.PrintHeaderFooter(detail_data[k].pos_hf, detail_data[k].pos_lcr, page_heder_footer[k]);
                        }

                    }
                }
                excel.SheetDelete(1);
                excel.SaveXls(path + character + "台本.xls");
                excel.CreatePDF(path + character + "台本.xls");
                excel.Close();
            }

        }
        // 緊急生成!
        string CreateVoiceFormat(int count, global.VoiceData voice)
        {
            return String.Format(voice.label + voice.middle + "{0:D" + voice.digit + "}", voice.serial + count);
        }
        // 緊急生成! key = string(キャラ名) value = global.VoiceData
        Hashtable label_hash_ = new Hashtable();
        void CreateLabelHash(List<global.VoiceData> voice_list)
        {
            Hashtable hoge = new Hashtable();
            foreach (var voice in voice_list)
            {
                string key = (string)overlap_label_[voice.key];
                if (!label_hash_.ContainsKey(key))
                    label_hash_[overlap_label_[voice.key]] = voice;
            }
        }
    }

    /// <summary>
    /// 台本化用クラス
    /// </summary>
    public partial class Libretto
    {
        class _Line
        {
            public string key;      // キャラ名
            public string serif;    // セリフ
            public int v_count;     // ボイスカウント
            public bool over_lap;   // 重複フラグ
            public bool cat_line;   // 改行フラグ
            public _Line(string k, string s, int cnt, bool lap, bool cat) { key = k; serif = s; v_count = cnt; over_lap = lap; cat_line = cat; }
        }
        class _Page
        {
            public List<_Line> lines = new List<_Line>();
            public void Add(string k, string s, int cnt, bool lap, bool cat) { lines.Add(new _Line(k, s, cnt, lap, cat)); }
            public void Add(_Line line) { lines.Add(new _Line(line.key, line.serif, line.v_count, line.over_lap, line.cat_line)); }
            public void AddRange(_Page page) { lines.AddRange(page.lines); }
            public int Length() { return lines.Count; }
            public _Line Index(int index) { return lines[index]; }
            public void Clear() { lines.Clear(); }
        }
        class _File
        {
            public List<_Page> pages = new List<_Page>();
            public string file_name = "";
            public void Add(_Page page) { pages.Add(page); }
        }
        class _Lib
        {
            public List<_File> files = new List<_File>();
            public void Add(_File file) { files.Add(file); }
        }
        class LibrettoKey
        {
            public int all_count;
            public string file_name;
            public int detail_count;
            public string s_serial = "";
            public string key = "";
            public string serif = "";
            public LibrettoKey(int all, string file, int detail, string serial, string k, string s)
            {
                all_count = all;
                file_name = file;
                detail_count = detail;
                s_serial = serial;
                key = k;
                serif = s;
            }
        }
        // 各キャラ毎のボイス番号をカウントしていく
        class MyHashKey
        {
            // key["キャラ名"] = 現在のボイスカウント++;
            public Hashtable key = new Hashtable();
            //id["0000_000000"] = key["キャラ名"]; (現在のボイスカウント)
            public Hashtable id = new Hashtable();
        };
        readonly int PAGE_TEXT_MAX = 30;
        readonly int PAGE_LINE_MAX = 35;
        _Lib lib_ = new _Lib();
        Hashtable overlap_label_ = new Hashtable(); // 重複したキーを同じキャラに設定
        Hashtable now_hash_ = new Hashtable();  // 重複を含めないラベル
        // 全体のボイス数をkey(キャラ)毎に与えていく
        MyHashKey hash_key_ = new MyHashKey();      // <string, int> キャラ名、カウント数
        bool comment_convert_ = false;
        // 台本を省略するかしないかのフラグ
        bool omission_ = false;
        // ページまたぎ時のセリフがつながっているかのフラグ
        bool page_straddle_ = false;
        

        /// <summary>
        /// これをやらないと、続けて台本化を何度か行った場合に前のデータが残ってまう。
        /// </summary>
        private void Initialize_()
        {
            lib_ = new _Lib();
            overlap_label_ = new Hashtable();
            now_hash_ = new Hashtable();
            hash_key_ = new MyHashKey();
            comment_convert_ = false;
        }
        private void CreateOverlapLabel_(List<global.VoiceData> voice_data)
        {
            // ラベル名が一緒の場合、同一のボイスとして扱う。
            Hashtable label_check = new Hashtable();
            foreach (var voice in voice_data)
            {
                // 新規登録
                if (!label_check.ContainsKey(voice.label))
                {
                    label_check[voice.label] = voice.key;
                    overlap_label_[voice.key] = voice.key;
                }

                // 重複
                else
                {
                    overlap_label_[voice.key] = label_check[voice.label];
                }
            }


            foreach (var value in label_check.Values)
            {
                now_hash_[value] = value;
            }
        }
        public void DebugLog(List<string> src_files, string dst_dir, List<global.VoiceData> voice_data)
        {
            Initialize_();
            // 関数の奥で使うため、毎回excel_data.conver_commentのためのbool引数を追加するのが面倒なので、ここで取得してます。
            comment_convert_ = false;
            CreateOverlapLabel_(voice_data);
            FileAdd_(src_files.ToArray(), false, "//");
            // debug出力
            var sb = new StringBuilder();
            foreach (var v in voice_data)
            {
                foreach (var file in lib_.files)
                {
                    foreach (var page in file.pages)
                    {
                        foreach (var line in page.lines)
                        {
                            if (line.v_count == -1) continue;
                            if(v.key == (string)overlap_label_[line.key])
                                sb.AppendLine("overlap[" + (string)overlap_label_[line.key] + "]" + line.v_count + ": " + "【" + v.key + "】" + line.serif);
                        }
                    }
                }
            }

            foreach (var key in hash_key_.key.Keys)
            {
                if (!now_hash_.ContainsKey(key)) continue;
                sb.AppendLine("総数: " + (string)key + " " + ((int)hash_key_.key[(string)key] + 1).ToString());
            }

            var file_name = dst_dir + "/debug.log";
            using (var sw = new StreamWriter(file_name, false, Encoding.Default))
            {
                sw.Write(sb.ToString());
            }
        }
        int GetKeyID(string line, int pos, ref bool over_lap)
        {
            if (pos >= PAGE_TEXT_MAX) return -1;

            if (comment_convert_)
                if(line.StartsWith("◇")) return - 1;
            else 
                if(line.StartsWith("//")) return -1;

            int serif_pos = line.IndexOf("「");

            if(serif_pos != -1){
                if (line.StartsWith("#!"))
                {
                    string now_id = line.Substring(2, 11);                  // 0000_000000
                    string n_now_key = line.Substring(13, serif_pos - 13);    // キャラ名
                    string now_key = (string)overlap_label_[n_now_key];
                    if (now_key == null) now_key = n_now_key;
                    if (!hash_key_.key.ContainsKey(now_key))
                    {
                        hash_key_.key[now_key] = 0;
                    }else{
                        hash_key_.key[now_key] = (int)hash_key_.key[now_key] + 1;
                    }
                    hash_key_.id[now_id] = hash_key_.key[now_key];
                    return (int)hash_key_.id[now_id];
                }
                else if (line.StartsWith("#?"))
                {
                    string now_id = line.Substring(2, 11);
                    over_lap = true;
                    return (int)hash_key_.id[now_id];
                }

                string m_now_key = line.Substring(0, serif_pos);
                string _now_key = (string)overlap_label_[m_now_key];
                if (_now_key == null) _now_key = m_now_key;
                if (!hash_key_.key.ContainsKey(_now_key))
                {
                    hash_key_.key[_now_key] = 0;
                }
                else
                {
                    hash_key_.key[_now_key] = (int)hash_key_.key[_now_key] + 1;
                }
                return (int)hash_key_.key[_now_key];
            }

            return -1;
        }
        void LineAdd_(string line, int pos, _Page page, _Page temp, bool voice)
        {
            uint cnt = (uint)((line.Length - pos) / PAGE_TEXT_MAX);
            // 重複判定
            bool over_lap = false;

            int v_count = GetKeyID(line, pos, ref over_lap);
            int starts_pos = 0;
            
            if (v_count != -1 && pos - 13 > 0)
            {
                starts_pos = 13;
                pos = pos - 13;
            }

            // 1行
            if (cnt == 0)
            {
                if (temp.Length() == 0)
                    //1行
                    temp.Add(line.Substring(starts_pos, pos), line.Substring(starts_pos + pos), v_count, over_lap, true);
                else
                    //loopの前が複数行だった
                    temp.Add("", line.Substring(starts_pos + pos), v_count, over_lap, true);
            }
            // 複数行
            else
            {
                if (temp.Length() == 0)
                    temp.Add(line.Substring(starts_pos, pos), line.Substring(starts_pos + pos, PAGE_TEXT_MAX), v_count, over_lap, false);
                else
                    temp.Add("", line.Substring(starts_pos + pos, PAGE_TEXT_MAX), v_count, over_lap, false);
                pos += starts_pos + PAGE_TEXT_MAX;
                //再帰処理
                LineAdd_(line, pos, page, temp, voice);
                return;
            }

            page.AddRange(temp);
        }
        void PageAdd_(int offset, _File file, _Page page)
        {
            var temp = new _Page();
            for (int i = 0; offset < page.Length(); ++i, ++offset)
            {
                if (i >= PAGE_LINE_MAX)
                {
                    file.Add(temp);
                    //再帰処理
                    PageAdd_(offset, file, page);
                    return;
                }
                temp.Add(page.Index(offset));
            }
            file.Add(temp);
        }
        void FileAdd_(string[] file_path, bool bool_comment, string comment)
        {

            // 各ファイルごと
            foreach (string path in file_path)
            {
                string[] lines = File.ReadAllLines(path, Encoding.Default);
                var page = new _Page();
                // 各ラインごと
                foreach (string l in lines)
                {
                    var temp = new _Page();
                    string line = l.Trim();

                    // コメントの場合変換
                    if (line.StartsWith(comment))
                    {
                        if (bool_comment)
                            line = line.Replace(comment, "◇");
                        LineAdd_(line, 0, page, temp, false);
                        continue;   // 処理を飛ばす
                    }

                    int serif_pos = line.IndexOf("「");
                    // セリフの場合
                    if (serif_pos != -1)
                    {
                        LineAdd_(line, serif_pos, page, temp, true);
                    }
                    else
                    {
                        int mind_voice_pos = line.IndexOf("（");
                        if (mind_voice_pos != -1)
                            LineAdd_(line, mind_voice_pos, page, temp, true);
                        else
                            LineAdd_(line, 0, page, temp, false);
                    }
                }

                var file = new _File();
                file.file_name = path;
                PageAdd_(0, file, page);
                lib_.Add(file);
            }
        }
        bool CheckKey_(string key)
        {
            if (now_hash_.ContainsKey(key)) return true;

            return false;
        }
        bool CreateLibrettoDirectorys_(string[] directorys)
        {
            foreach (var directory in directorys)
            {
                if (!util.Util.CreateDirectory(directory)) return false;
            }
            return true;
        }
        bool PageKeyCheck_(List<_Page> pages, string key)
        {
            foreach (var page in pages)
            {
                foreach (var line in page.lines)
                {
                    if (key == (string)overlap_label_[line.key])
                        if(!line.over_lap)
                            return true;
                }
            }
            return false;
        }
        bool CreateVoiceExcel_(MyExcel excel, List<global.VoiceData> voice_data, global.ExcelData excel_data, string directory)
        {
            foreach (var v in voice_data)
            {
                // 最初に、被っているラベルの場合は台本を作らなくてよいのでその処理
                //if (!now_hash_.ContainsKey(v.key)) continue;
                if (!CheckKey_(v.key)) continue;
                excel.CreateBook(1);
                excel.SetZoom(80);
                var lk = new List<LibrettoKey>();
                int all_count = 0;
                int page_count = 2;

                foreach (var file in lib_.files)
                {
                    string file_name = Path.GetFileNameWithoutExtension(file.file_name);
                    int str_cat_index = -1;

                    if (CheckFileKey(file, v.key)  == 0) continue;

                    if (!PageKeyCheck_(file.pages, v.key)) continue;


                    foreach (var page in file.pages)
                    {
                        // ページまたぎ防止用
                        if (str_cat_index == -1)
                            if (CheckPageKey(page, v.key) == 0) continue;

                        ++page_count;
                        foreach (var line in page.lines)
                        {
                            if (v.key == (string)overlap_label_[line.key])
                            {
                                if (line.over_lap) continue;

                                ++all_count;
                                var s_serial = String.Format(v.label + v.middle + "{0:D" + v.digit + "}", (v.serial + line.v_count));  //(serial_count++)));
                                s_serial = v.top + s_serial + v.bottom;
                                lk.Add(new LibrettoKey(all_count, file_name, page_count, s_serial, line.key, line.serif));
                                if (line.serif.IndexOf("」") != -1 || line.serif.IndexOf("）") != -1)
                                    str_cat_index = -1;
                                else
                                    str_cat_index = lk.Count - 1;
                            }
                            else if (str_cat_index != -1)
                            {
                                lk[str_cat_index].serif += line.serif;
                                if (line.serif.IndexOf("」") != -1 || line.serif.IndexOf("）") != -1) str_cat_index = -1;
                            }
                        }
                    }
                    ++page_count;   //ファイル名のページ分足す。
                }
                //rui
                object[,] cell = new object[lk.Count + 1, 7];
                cell[0, 0] = "収録数";
                cell[0, 1] = "シーン名";
                cell[0, 2] = "台本項";
                cell[0, 3] = "音声番号";
                cell[0, 4] = "備考";
                cell[0, 5] = "台詞(収録後)";
                cell[0, 6] = "台詞(収録前)";
                for (int i = 0; i < lk.Count; ++i)
                {
                    cell[i + 1, 0] = lk[i].all_count;
                    cell[i + 1, 1] = lk[i].file_name;
                    cell[i + 1, 2] = lk[i].detail_count;
                    cell[i + 1, 3] = lk[i].s_serial;
                    //cell[i + 1, 4] = " ";
                    cell[i + 1, 5] = lk[i].serif;
                    cell[i + 1, 6] = cell[i + 1, 5];
                }
                excel.CreateCell(0, 0, cell);
                excel.SetRange();
                excel.SetFontColor(MyExcel.EXCEL_MATRIX.X, 0, Color.White);
                excel.SetInteriorColor(MyExcel.EXCEL_MATRIX.X, 0, Color.Green);
                excel.SetFontName(excel_data.font_name);    //デフォルトフォントなら処理しない方が断然早い
                //excel.SetWidthAutoFit();

                excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.Y, 0, 7);
                excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.Y, 1, 25);
                excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.Y, 2, 7);
                excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.Y, 3, 17);
                excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.Y, 4, 7);
                excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.Y, 5, 64);
                excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.Y, 6, 64);

                // 条件付き書式 値変更時の色
                excel.ConditionNotEqual(MyExcel.EXCEL_MATRIX.X, 5, 6, Color.PowderBlue);
                // 条件付き書式 同じ文字列の色
                //excel.ConditionOverlap(MyExcel.EXCEL_MATRIX.X, 6, Color.HotPink);
                excel.SetBorderColor(Color.Navy);
                excel.CellWrapMode();
                //PDF作成
                excel.PrintOrientation();
                excel.CreatePDF(directory + v.key + "_抜き台本.xlsx", 0, 5);

                //ブック毎にセーブ xlsx作成
                excel.SaveXlsx(directory + v.key + "_抜き台本.xlsx");
                excel.Close();
            }

#if false
            try
            {

                foreach (var voice in voice_data)
                {
                    // 最初に、被っているラベルの場合は台本を作らなくてよいのでその処理
                    if (!now_hash_.ContainsKey(overlap_label_[voice.key])) continue;

                    excel.CreateBook(1);
                    excel.SetZoom(80);
                    var lk = new List<LibrettoKey>();
                    int all_count = 0;
                    int serial_count = 0;
                    int page_count = 2;

                    //キー作成
                    foreach (var lib_file in lib_.files)
                    {
                        string file_name = Path.GetFileNameWithoutExtension(lib_file.file_name);
                        int str_cat_index = -1;

                        //if (!CheckKey(lib_file.pages, (string)overlap_label_[voice.key])) continue;
                        if (!CheckKey(lib_file.pages, voice.key)) continue;

                        foreach (var file in lib_file.pages)
                        {
                            ++page_count;
                            for(int i = 0; i < file.lines.Count; ++i)
                            {
                                // test
                                if(voice.key == overlap_label_[file.lines[i].key])
                                {
                                    ++all_count;
                                    var s_serial = String.Format(voice.label + voice.middle + "{0:D" + voice.digit + "}", (voice.serial + file.lines[i].v_count));  //(serial_count++)));
                                    s_serial = voice.top + s_serial + voice.bottom;
                                    lk.Add(new LibrettoKey(all_count, file_name, page_count, s_serial, file.lines[i].key, file.lines[i].serif));
                                    if (file.lines[i].serif.IndexOf("」") != -1 || file.lines[i].serif.IndexOf("）") != -1) 
                                        str_cat_index = -1;
                                    else
                                        str_cat_index = lk.Count - 1;

                                }
                                else if (str_cat_index != -1)
                                {
                                    lk[str_cat_index].serif += file.lines[i].serif;
                                    if (file.lines[i].serif.IndexOf("」") != -1 || file.lines[i].serif.IndexOf("）") != -1) str_cat_index = -1;
                                }
                            }
                        }
                        ++page_count;   //ファイル名のページ分足す。
                    }

                    object[,] cell = new object[lk.Count + 1, 7];
                    cell[0, 0] = "収録数";
                    cell[0, 1] = "シーン名";
                    cell[0, 2] = "台本項";
                    cell[0, 3] = "音声番号";
                    cell[0, 4] = "備考";
                    cell[0, 5] = "台詞(収録後)";
                    cell[0, 6] = "台詞(収録前)";
                    for (int i = 0; i < lk.Count; ++i)
                    {
                        cell[i + 1, 0] = lk[i].all_count;
                        cell[i + 1, 1] = lk[i].file_name;
                        cell[i + 1, 2] = lk[i].detail_count;
                        cell[i + 1, 3] = lk[i].s_serial;
                        //cell[i + 1, 4] = " ";
                        cell[i + 1, 5] = lk[i].serif;
                        cell[i + 1, 6] = cell[i + 1, 5];
                    }
                    excel.CreateCell(0, 0, cell);
                    excel.SetRange();
                    excel.SetFontColor(MyExcel.EXCEL_MATRIX.X, 0, Color.White);
                    excel.SetInteriorColor(MyExcel.EXCEL_MATRIX.X, 0, Color.Green);
                    excel.SetFontName(excel_data.font_name);    //デフォルトフォントなら処理しない方が断然早い
                    //excel.SetWidthAutoFit();

                    excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.Y, 0, 7);
                    excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.Y, 1, 25);
                    excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.Y, 2, 7);
                    excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.Y, 3, 17);
                    excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.Y, 4, 7);
                    excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.Y, 5, 64);
                    excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.Y, 6, 64);

                    // 条件付き書式 値変更時の色
                    excel.ConditionNotEqual(MyExcel.EXCEL_MATRIX.X, 5, 6, Color.PowderBlue);
                    // 条件付き書式 同じ文字列の色
                    //excel.ConditionOverlap(MyExcel.EXCEL_MATRIX.X, 6, Color.HotPink);
                    excel.SetBorderColor(Color.Navy);
                    excel.CellWrapMode();
                    //PDF作成
                    excel.PrintOrientation();
                    excel.CreatePDF(directory + voice.key + "_抜き台本.xlsx", 0, 5);

                    //ブック毎にセーブ xlsx作成
                    excel.SaveXlsx(directory + voice.key + "_抜き台本.xlsx");
                    excel.Close();
                }
            }
            catch
            {
                Debug.WriteLine("error発生! : CreateVoiceExcel_");
                excel.End();
                return false;
            }
#endif
            return true;
        }

        bool CreateLibrettoExcelAndPDF_(MyExcel excel, List<global.VoiceData> voice_data, global.ExcelData excel_data, global.HeadFootPos[] detail_data, string[] directorys, BackgroundWorker worker)
        {
            // プログレスバー用の%を出す 最初で5%使用するので残り95%での割合
            // 1ボイス終了するごとにperを足す
            int per = (int)((float)(9.5f / now_hash_.Count) * 10.0f); // 途中でmaxなっちゃうかもなので-3で回避

            try
            {
                foreach (var voice in voice_data)
                {
                    // 最初に、被っているラベルの場合は台本を作らなくてよいのでその処理
                    if (!now_hash_.ContainsKey(voice.key)) continue;

                    excel.CreateBook(1);
                    int page_count = 0;
                    //先頭のアンダーバーはここ以外で使うとバグる可能性があるprivateの中のprivateだと思ってください。 (つまりここ以外では利用不可です)
                    //表紙作成 rui
                    _Cover(excel, voice, excel_data, detail_data, ref page_count);

                    // プログレスバー処理
                    worker.ReportProgress(per, "キャラ名「" + voice.key + "」進行中");

                    _ContentsExcelAndPDF(excel, voice, excel_data, detail_data, ref page_count);
                    excel.SheetDelete(1);
                    excel.SaveXls(directorys[0] + voice.key + "_台本.xls");
                    excel.CreatePDF(directorys[1] + voice.key + ".xls");
                    excel.Close();

                    // プログレスバー処理 表示のみ
                    worker.ReportProgress(0, "キャラ名「" + voice.key + "」終了");
                }
            }
            catch(Exception e)
            {
                MessageBox.Show(e.ToString(), "エラー発生");
                Debug.WriteLine("エラー発生" + e);
                excel.End();
                return false;
            }
            return true;
        }
        void _Cover(MyExcel excel, global.VoiceData voice, global.ExcelData excel_data, global.HeadFootPos[] detail_data, ref int page_count)
        {
            int page_max = OmissionFileMax(lib_, voice.key);

            //セル設定
            //excel.CreateNewSheet(voice.key + (++page_count).ToString());
            excel.CreateNewSheet((++page_count).ToString());
            excel.CreateCell(0, 0, 3, 1);

            excel.SetCell(0, 0, excel_data.title);
            excel.SetCell(1, 0, "バージョン : " + excel_data.version);
            excel.SetCell(2, 0, "日付 : " + excel_data.today_date);

            excel.SetRange();
            excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.X, 0, 50);
            excel.CellAlignment(MyExcel.CELL_H_DIR.CENTER, MyExcel.EXCEL_MATRIX.Y, 0);
            excel.SetFontSize(26);
            excel.SetWidthAutoFit();
            excel.SetHeightAutoFit();

            //印刷設定
            excel.PrintMiddle();
            excel.PrintOrientation();
            excel.PrintHeaderFooter(detail_data[1].pos_hf, detail_data[1].pos_lcr, page_count.ToString() + "/" + page_max.ToString());
            excel.PrintSize(MyExcel.PRINT_SIZE.A4);
        }
        
        /*
        int KeyPageMax_(_Lib lib, string key)
        {
            int cnt = 1;
            // 各ファイル毎
            foreach (var lib_file in lib_.files)
            {
                if (!PageKeyCheck_(lib_file.pages, key)) continue;
                cnt += lib_file.pages.Count + 1;
            }
            return cnt;
        }*/
        int CheckFileKey(_File file, string key)
        {
            int key_count = 0;
            foreach (var page in file.pages)
            {
                key_count += CheckPageKey(page, key);
            }

            return key_count;
        }

        int OmissionFileMax(_Lib lib)
        {
            return AllBooks(lib);
        }
        int OmissionFileMax(_Lib lib, string key)
        {
            if (omission_)
                return TEST(lib_, key);

            return AllBooks(lib);
        }
        int AllBooks(_Lib lib)
        {
            int page_count = 1;
            foreach (var file in lib.files)
            {
                 page_count += file.pages.Count + 1;
            }

            return page_count;
        }
        int TEST(_Lib lib, string key)
        {
            int page_count = 1;

            foreach (var file in lib.files)
            {
                int cat = -1;
                if (CheckFileKey(file, key) == 0) continue;
                if (!PageKeyCheck_(file.pages, key)) continue;
                foreach (var page in file.pages)
                {
                    if(cat == -1)
                        if (CheckPageKey(page, key) == 0) continue;
                    ++page_count;

                    foreach (var line in page.lines)
                    {
                        if (key == (string)overlap_label_[line.key])
                        {
                            if (line.over_lap) continue;
                            if (line.serif.IndexOf("」") != -1 || line.serif.IndexOf("）") != -1)
                                cat = -1;
                            else
                                cat = 0;
                        }
                        else if (cat != -1)
                        {
                            if (line.serif.IndexOf("」") != -1 || line.serif.IndexOf("）") != -1) cat = -1;
                        }
                    }
                }
                // ファイル表紙分
                ++page_count;
            }

            return page_count;
        }

        int DebugCheckPageKey(_Page page, string key, string file_name)
        {
            int key_count = 0;
            int over_lap = 0;
            // ページまたぎ防止 page_straddle_
            foreach (var line in page.lines)
            {
                if ((string)overlap_label_[line.key] == key)
                {
                    if (line.over_lap)
                        ++over_lap;
                    else
                    {
                        ++key_count;
                    }
                    if (!line.cat_line) page_straddle_ = true;
                }
                else if (line.cat_line) page_straddle_ = false;
            }
            // この状態の場合はページを作成したいので書き込む
            if ((over_lap + key_count) == 0) return 1;

            return key_count;
        }
        int CheckPageKey(_Page page, string key)
        {
            int key_count = 0;
            int over_lap = 0;
            // ページまたぎ防止 page_straddle_

            foreach (var line in page.lines)
            {
                if ((string)overlap_label_[line.key] == key)
                {
                    if (line.over_lap)
                        ++over_lap;
                    else
                        ++key_count;

                    if (!line.cat_line) page_straddle_ = true;
                }
                else if (line.cat_line) page_straddle_ = false;
            }

            // この状態の場合はページを作成したいので書き込む
            if ((over_lap + key_count) == 0) return 1;

            return key_count;
        }
        //int CheckMax(_Lib lib, string key)
        //{
        //    int all_count = 1;  // 最初の表紙

        //    foreach (var file in lib.files)
        //    {
        //        // セリフが１つもない場合飛ばす
        //        if (!PageKeyCheck_(file.pages, key)) continue;

        //        page_straddle_ = false;
        //        int page_count = 0;
                

        //        foreach (var page in file.pages)
        //        {
        //            if (page_straddle_){
        //                page_straddle_ = false;
        //                ++page_count;
        //            }else{
        //                //if (CheckPageKey(page, key) != 0)
        //                //{
        //                //    ++page_count;
        //                //}
        //                if (DebugCheckPageKey(page, key, file.file_name) != 0)
        //                    ++page_count;

        //            }
        //        }

        //        if (page_count != 0)
        //        {
        //            all_count += page_count + 1;    // ファイル毎の表紙の+1
        //        }
        //    }
           
        //    return all_count;
        //}
        void _ContentsExcelAndPDF(MyExcel excel, global.VoiceData voice, global.ExcelData excel_data, global.HeadFootPos[] detail_data, ref int page_count)
        {

            int page_max = OmissionFileMax(lib_, voice.key);

            int voice_count = 0;
            //int serial_count = 0;

            // 各ファイル毎
            foreach (var lib_file in lib_.files)
            {
                if (omission_)
                {
                    // 全チェック用
                    if (CheckFileKey(lib_file, voice.key) == 0) continue;

                    if (!PageKeyCheck_(lib_file.pages, voice.key)) continue;
                }
                excel.CreateNewSheet((++voice_count + 1).ToString());   //カバー分の + 1
                //excel.CreateNewSheet(voice.key + (++voice_count + 1).ToString());   //カバー分の + 1
                excel.CreateCell(0, 0, 2, 1);
                excel.SetCell(0, 0, "シーン : " + Path.GetFileNameWithoutExtension(lib_file.file_name));
                excel.SetCell(1, 0, voice.key + "台本");
                excel.SetRange();
                excel.SetFontSize(22);

                //印刷設定
                excel.CellAlignment(MyExcel.CELL_H_DIR.CENTER, MyExcel.EXCEL_MATRIX.Y, 0);
                excel.PrintMiddle();
                excel.PrintOrientation();
                excel.SetWidthAutoFit();
                
                excel.PrintHeaderFooter(detail_data[1].pos_hf, detail_data[1].pos_lcr, (++page_count).ToString() + "/" + page_max);
                excel.PrintSize(MyExcel.PRINT_SIZE.A4);

                bool str_cat = false;
                bool over_lap_cat = false;

                // 各ページ毎
                foreach (var page in lib_file.pages)
                {
                    if (omission_)
                    {
                        // ページマタギ防止用
                        if (!str_cat)
                            if (CheckPageKey(page, voice.key) == 0) continue;
                    }

                    //if (CheckPageOverlap(page, voice.key)) continue;

                    //if (voice.key == overlap_label_[page.lines.key])
                    excel.CreateNewSheet((++voice_count + 1).ToString());   //カバー分の + 1
                    //excel.CreateNewSheet(voice.key + (++voice_count + 1).ToString());   //カバー分の + 1

                    excel.CreateCell(0, 0, PAGE_LINE_MAX + 1, 3);
                    excel.SetCell(0, 0, "音声番号");
                    excel.SetCell(0, 1, "キャラ名");
                    excel.SetCell(0, 2, "セリフ");
                    List<int> voice_index = new List<int>();
                    List<int> over_lap_index = new List<int>();
                    int i = 0;
                    foreach (var line in page.lines)
                    {
                        // 各キャラ毎
                        if ((string)overlap_label_[line.key] == voice.key)
                        {
                            var s_serial = String.Format(voice.label + voice.middle + "{0:D" + voice.digit + "}", (voice.serial + line.v_count));//(serial_count++)));
                            // ボイス番号セット
                            excel.SetCell(i + 1, 0, s_serial);
                            // キャラ名セット ※スターフラグが付いている場合、☆追加
                            if (excel_data.check_star)
                                excel.SetCell(i + 1, 1, "☆" + line.key);
                            else
                                excel.SetCell(i + 1, 1, line.key);

                            if (line.over_lap)
                            {
                                if (line.serif.IndexOf("」") != -1)
                                    over_lap_cat = false;
                                else
                                    over_lap_cat = true;

                                over_lap_index.Add(i + 1);
                            }
                            else
                            {
                                if (line.serif.IndexOf("」") != -1)
                                    str_cat = false;
                                else
                                    str_cat = true;
                                voice_index.Add(i + 1);
                            }
                        }
                        else if (str_cat)
                        {
                            voice_index.Add(i + 1);
                            if (line.serif.IndexOf("」") != -1)
                                str_cat = false;
                        }
                        else if (over_lap_cat)
                        {
                            over_lap_index.Add(i + 1);
                            if (line.serif.IndexOf("」") != -1)
                                over_lap_cat = false;
                        }
                        else
                        {
                            excel.SetCell(i + 1, 1, line.key);
                        }
                        excel.SetCell(i + 1, 2, line.serif);

                        ++i;
                    }
                    excel.ChangeXY();
                    excel.SetRange();


                    // 全体のフォント色
                    SetFontColor_(excel, excel_data.other_voice_color);

                    // カブリボイス色 灰色固定
                    SetFontColor_(excel, over_lap_index, Color.Gray);

                    // ボイス色 ※太字なら太字追加
                    if (excel_data.bold)
                        SetFontBold_(excel, voice_index, excel_data.voice_color);
                    else
                        SetFontColor_(excel, voice_index, excel_data.voice_color);

                    excel.SellxlVertical();                     //縦書き設定
                    excel.SetFontName(excel_data.font_name);    //フォント設定

                    //ボーダー設定
                    Color border_color = Color.Green;
                    excel.BorderRange(MyExcel.EXCEL_MATRIX.Y, excel.EndX() - 1, border_color);
                    excel.BorderRange(MyExcel.BORDER_RANGE.EDGE_BOTTOM, border_color);
                    excel.BorderRange(MyExcel.BORDER_RANGE.EDGE_LEFT, border_color);
                    //excel.BorderRange(MyExcel.BORDER_RANGE.EDGE_RIGHT);
                    excel.BorderRange(MyExcel.BORDER_RANGE.EDGE_TOP, border_color);
                    excel.BorderRange(MyExcel.BORDER_RANGE.INSIDE_HORIZONTAL, border_color);


                    //excel.SellxOrientation(-90, MyExcel.EXCEL_MATRIX.X, 0);                           
                    excel.SellxOrientation(-90, MyExcel.EXCEL_MATRIX.X, 0, excel.StartX(), excel.EndX() - 2);  //ボイス行のみ横文字、横表示
                    excel.CellAlignment(MyExcel.CELL_H_DIR.CENTER, MyExcel.EXCEL_MATRIX.X, 0);        //ボイス行のみ、文字詰めセンター調整

                    //各種セル手動調整
                    excel.SetFontSize(excel_data.font_size);
                    excel.CellAlignment(MyExcel.CELL_V_DIR.TOP);
                    excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.X, 0, 3);
                    excel.CellSizeHeight(MyExcel.EXCEL_MATRIX.X, 0, 70);
                    excel.CellSizeHeight(MyExcel.EXCEL_MATRIX.X, 1, 80);
                    excel.CellSizeHeight(MyExcel.EXCEL_MATRIX.X, 2, 385);

                    //印刷設定
                    excel.PrintOrientation();
                    excel.PrintArea();
                    var s = new string[] { 
                        detail_data[0].print,                                       //ロゴ
                        (++page_count).ToString() + "/" + page_max,                 //ページ数
                        Path.GetFileNameWithoutExtension(lib_file.file_name),       //ファイル名
                        voice.key + "台本",                                         //台本名
                        DateTime.Today.ToShortDateString()                          //日付
                    };

                    for (int k = 0; k < s.Length; k++)
                    {
                        excel.PrintHeaderFooter(detail_data[k].pos_hf, detail_data[k].pos_lcr, s[k]);
                    }

                    excel.PrintSize(MyExcel.PRINT_SIZE.A4);
                }
            }
        }
        void SetFontColor_(MyExcel excel, Color color)
        {
            excel.SetFontColor(color);
        }

        void SetFontColor_(MyExcel excel, List<int> index_list, Color color)
        {
            foreach (var index in index_list)
            {
                excel.SetBold(MyExcel.EXCEL_MATRIX.Y, PAGE_LINE_MAX - index, color);
            }
        }
        void SetFontBold_(MyExcel excel, List<int> index_list, Color color)
        {
            foreach (var index in index_list)
            {
                excel.SetBold(MyExcel.EXCEL_MATRIX.Y, PAGE_LINE_MAX - index, color);
            }
        }
        void CreateConvertText_(List<global.VoiceData> voice_data, string directory)
        {
            foreach (var lib_file in lib_.files)
            {
                var sb = new StringBuilder();
                foreach (var page in lib_file.pages)
                {
                    foreach (var line in page.lines)
                    {
                        int pos = line.serif.IndexOf("「");
                        int end_pos = line.serif.IndexOf("」");
                        if (pos != -1)
                        {
                            foreach (var data in voice_data)
                            {
                                if (data.key == line.key)
                                {
                                    var temp = String.Format(data.label + data.middle + "{0:D" + data.digit + "}", (data.serial + line.v_count));
                                    var voice_format = data.top + temp + data.bottom;
                                    sb.AppendLine(voice_format);
                                    if(line.cat_line)
                                        sb.AppendLine(line.key + line.serif);
                                    else
                                        sb.Append(line.key + line.serif);
                                    goto hoge1;
                                }
                            }
                            if (line.cat_line)
                                sb.AppendLine(line.key + line.serif);
                            else
                                sb.Append(line.key + line.serif);
                            hoge1: ;
                        }
                        else
                        {
                            int pos2 = line.serif.IndexOf("（");
                            if (pos2 != -1)
                            {
                                foreach (var data in voice_data)
                                {
                                    if (data.key == line.key)
                                    {
                                        var temp = String.Format(data.label + data.middle + "{0:D" + data.digit + "}", (data.serial + line.v_count));
                                        var voice_format = data.top + temp + data.bottom;
                                        sb.AppendLine(voice_format);
                                        if(line.cat_line)
                                            sb.AppendLine(line.key + line.serif);
                                        else
                                            sb.Append(line.key + line.serif);
                                        goto hoge2;
                                    }
                                }
                                if (line.cat_line)
                                    sb.AppendLine(line.key + line.serif);
                                else
                                    sb.Append(line.key + line.serif);
                                hoge2: ;
                            }
                            else
                            {
                                if(line.cat_line)
                                    sb.AppendLine(line.serif);
                                else
                                    sb.Append(line.serif);
                            }
                        }
                    }
                }

                var file_name = Path.GetFileName(lib_file.file_name);
                using (var sw = new StreamWriter(directory + file_name, false, Encoding.Default))
                {
                    sw.Write(sb.ToString());
                }

            }
        }
        void CreateBackUpText_(string src_file, string dst_dir)
        {
            var file_BAK = Path.GetFileName(src_file);
            file_BAK = Path.ChangeExtension(file_BAK, ".BAK");
            util.Util.FileCopy(src_file, dst_dir + file_BAK);

        }

        void test_(MyExcel excel, global.ExcelData excel_data, global.HeadFootPos[] detail_data, ref int page_count)
        {

            int page_max = OmissionFileMax(lib_);

            int voice_count = 0;

            // 各ファイル毎
            foreach (var lib_file in lib_.files)
            {
                excel.CreateNewSheet((++voice_count + 1).ToString());   //カバー分の + 1
                excel.CreateCell(0, 0, 2, 1);
                excel.SetCell(0, 0, "シーン : " + Path.GetFileNameWithoutExtension(lib_file.file_name));
                excel.SetCell(1, 0, "台本");
                excel.SetRange();
                excel.SetFontSize(22);

                //印刷設定
                excel.CellAlignment(MyExcel.CELL_H_DIR.CENTER, MyExcel.EXCEL_MATRIX.Y, 0);
                excel.PrintMiddle();
                excel.PrintOrientation();
                excel.SetWidthAutoFit();

                excel.PrintHeaderFooter(detail_data[1].pos_hf, detail_data[1].pos_lcr, (++page_count).ToString() + "/" + page_max);
                excel.PrintSize(MyExcel.PRINT_SIZE.A4);


                // 各ページ毎
                foreach (var page in lib_file.pages)
                {
                    excel.CreateNewSheet((++voice_count + 1).ToString());   //カバー分の + 1
                    excel.CreateCell(0, 0, PAGE_LINE_MAX + 1, 3);
                    excel.SetCell(0, 0, "音声番号");
                    excel.SetCell(0, 1, "キャラ名");
                    excel.SetCell(0, 2, "セリフ");
                    List<int> voice_index = new List<int>();
                    List<int> over_lap_index = new List<int>();
                    int i = 0;
                    foreach (var line in page.lines)
                    {
                        excel.SetCell(i + 1, 1, line.key);
                        excel.SetCell(i + 1, 2, line.serif);
                        ++i;
                    }
                    excel.ChangeXY();
                    excel.SetRange();

                    excel.SellxlVertical();                     //縦書き設定
                    excel.SetFontName(excel_data.font_name);    //フォント設定

                    //ボーダー設定
                    Color border_color = Color.Green;
                    excel.BorderRange(MyExcel.EXCEL_MATRIX.Y, excel.EndX() - 1, border_color);
                    excel.BorderRange(MyExcel.BORDER_RANGE.EDGE_BOTTOM, border_color);
                    excel.BorderRange(MyExcel.BORDER_RANGE.EDGE_LEFT, border_color);
                    //excel.BorderRange(MyExcel.BORDER_RANGE.EDGE_RIGHT);
                    excel.BorderRange(MyExcel.BORDER_RANGE.EDGE_TOP, border_color);
                    excel.BorderRange(MyExcel.BORDER_RANGE.INSIDE_HORIZONTAL, border_color);


                    //excel.SellxOrientation(-90, MyExcel.EXCEL_MATRIX.X, 0);                           
                    excel.SellxOrientation(-90, MyExcel.EXCEL_MATRIX.X, 0, excel.StartX(), excel.EndX() - 2);  //ボイス行のみ横文字、横表示
                    excel.CellAlignment(MyExcel.CELL_H_DIR.CENTER, MyExcel.EXCEL_MATRIX.X, 0);        //ボイス行のみ、文字詰めセンター調整

                    //各種セル手動調整
                    excel.SetFontSize(excel_data.font_size);
                    excel.CellAlignment(MyExcel.CELL_V_DIR.TOP);
                    excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.X, 0, 3);
                    excel.CellSizeHeight(MyExcel.EXCEL_MATRIX.X, 0, 70);
                    excel.CellSizeHeight(MyExcel.EXCEL_MATRIX.X, 1, 80);
                    excel.CellSizeHeight(MyExcel.EXCEL_MATRIX.X, 2, 385);

                    //印刷設定
                    excel.PrintOrientation();
                    excel.PrintArea();
                    var s = new string[] { 
                        detail_data[0].print,                                       //ロゴ
                        (++page_count).ToString() + "/" + page_max,                 //ページ数
                        Path.GetFileNameWithoutExtension(lib_file.file_name),       //ファイル名
                        "台本",                                                     //台本名
                        DateTime.Today.ToShortDateString()                          //日付
                    };

                    for (int k = 0; k < s.Length; k++)
                    {
                        excel.PrintHeaderFooter(detail_data[k].pos_hf, detail_data[k].pos_lcr, s[k]);
                    }

                    excel.PrintSize(MyExcel.PRINT_SIZE.A4);
                }
            }
        }
        void _testCover(MyExcel excel, global.ExcelData excel_data, global.HeadFootPos[] detail_data, ref int page_count)
        {
            int page_max = OmissionFileMax(lib_);

            //セル設定
            //excel.CreateNewSheet(voice.key + (++page_count).ToString());
            excel.CreateNewSheet((++page_count).ToString());
            excel.CreateCell(0, 0, 3, 1);

            excel.SetCell(0, 0, excel_data.title);
            excel.SetCell(1, 0, "バージョン : " + excel_data.version);
            excel.SetCell(2, 0, "日付 : " + excel_data.today_date);

            excel.SetRange();
            excel.CellSizeWidth(MyExcel.EXCEL_MATRIX.X, 0, 50);
            excel.CellAlignment(MyExcel.CELL_H_DIR.CENTER, MyExcel.EXCEL_MATRIX.Y, 0);
            excel.SetFontSize(26);
            excel.SetWidthAutoFit();
            excel.SetHeightAutoFit();

            //印刷設定
            excel.PrintMiddle();
            excel.PrintOrientation();
            excel.PrintHeaderFooter(detail_data[1].pos_hf, detail_data[1].pos_lcr, page_count.ToString() + "/" + page_max.ToString());
            excel.PrintSize(MyExcel.PRINT_SIZE.A4);
        }
        bool test2_(MyExcel excel, List<global.VoiceData> voice_data, global.ExcelData excel_data, global.HeadFootPos[] detail_data, string[] directorys, BackgroundWorker worker)
        {
            excel.CreateBook(1);
            int page_count = 0;
            _testCover(excel, excel_data, detail_data, ref page_count);
            test_(excel, excel_data, detail_data, ref page_count);
            excel.SheetDelete(1);
            excel.SaveXls(directorys[0] + "台本.xls");
            excel.CreatePDF(directorys[1] + "台本.xls");
            excel.Close();

            return true;
        }
    }
}
