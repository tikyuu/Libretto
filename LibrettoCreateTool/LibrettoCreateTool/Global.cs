using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;

//台本作成用のグローバルクラス
namespace global
{
    public struct OutputFlags
    {

        public bool libretto_omission;   // 台本省略型
        public bool voice_excel;
        public bool libretto;
        public bool lib_only;
        public bool voice_lib_only;
        public OutputFlags(bool omission, bool voice, bool lib, bool only, bool v_lib_only)
        {
            libretto_omission = omission;
            voice_excel  = voice;
            libretto = lib;
            lib_only = only;
            voice_lib_only = v_lib_only;
        }
    }
    /// <summary>
    /// 作成するキャラのデータ
    /// </summary>
    public class VoiceData
    {
        /// <summary>
        /// キャラ名
        /// </summary>
        public string key;
        /// <summary>
        /// フォーマット名
        /// </summary>
        public string label;
        /// <summary>
        /// シリアル番号
        /// </summary>
        public int serial;
        /// <summary>
        /// 桁数
        /// </summary>
        public int digit;
        /// <summary>
        /// 【
        /// </summary>
        public string top;
        /// <summary>
        /// _
        /// </summary>
        public string middle;
        /// <summary>
        /// 】
        /// </summary>
        public string bottom;
        /// <summary>
        /// 通し台本作成フラグ
        /// </summary>
        public bool serial_lib;
        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="_key">キャラ名</param>
        /// <param name="_label">フォーマット名</param>
        /// <param name="_serial">シリアル番号</param>
        /// <param name="_digit">桁数</param>
        /// <param name="_top">【</param>
        /// <param name="_middle">_</param>
        /// <param name="_bottom">】</param>
        public VoiceData(string _key, string _label, int _serial, int _digit, string _top, string _middle, string _bottom, bool _serial_lib)
        {
            key = _key;
            label = _label;
            serial = _serial;
            digit = _digit;
            top = _top;
            middle = _middle;
            bottom = _bottom;
            serial_lib = _serial_lib;
        }
    }
    /// <summary>
    /// PDF作成のヘッダーフッター専用のクラス
    /// </summary>
    public class HeadFootPos
    {
        /// <summary>
        /// h = ヘッダー f = フッター
        /// </summary>
        public int pos_hf;
        /// <summary>
        /// l = レフト c = センター r ライト
        /// </summary>
        public int pos_lcr;
        /// <summary>
        /// 表示する文字列
        /// </summary>
        public string print = "";

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="hf">h = ヘッダー f = フッター</param>
        /// <param name="lcr">l = レフト c = センター r ライト</param>
        /// <param name="p">表示する文字列</param>
        public HeadFootPos(int hf, int lcr, string p)
        {
            pos_hf = hf;
            pos_lcr = lcr;
            print = p;
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="hf">h = ヘッダー f = フッター</param>
        /// <param name="lcr">l = レフト c = センター r ライト</param>
        public HeadFootPos(int hf, int lcr)
        {
            pos_hf = hf;
            pos_lcr = lcr;
        }
    }
    /// <summary>
    /// エクセル作成用の設定データ
    /// </summary>
    public class ExcelData
    {
        /// <summary>
        /// ズーム
        /// </summary>
        public int zoom;
        /// <summary>
        /// フォントサイズ
        /// </summary>
        public int font_size;
        /// <summary>
        /// コメント
        /// </summary>
        public string comment;
        /// <summary>
        /// タイトル
        /// </summary>
        public string title;
        /// <summary>
        /// 太字
        /// </summary>
        public bool bold = false;
        /// <summary>
        /// ボイスキャラ以外のボイスカラー
        /// </summary>
        public Color other_voice_color = Color.Black;   
        /// <summary>
        /// ボイスキャラのカラー
        /// </summary>
        public Color voice_color = Color.Black;   
        /// <summary>
        /// 指定フォント
        /// </summary>
        public string font_name;
        /// <summary>
        /// 折り返し最大文字数
        /// </summary>
        public int wrap_max;
        /// <summary>
        /// バージョン管理
        /// </summary>
        public string version;
        /// <summary>
        /// 日付
        /// </summary>
        public string today_date;
        /// <summary>
        /// コメント変換判定
        /// </summary>
        public bool convert_comment;
        /// <summary>
        /// ☆追加判定
        /// </summary>
        public bool check_star;
        /// <summary>
        /// エクセル可視判定
        /// </summary>
        public bool excel_visible;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="zoom">ズーム値</param>
        /// <param name="font_size">フォントサイズ</param>
        /// <param name="comment">コメント文字</param>
        /// <param name="title">タイトル</param>
        /// <param name="bold">太字チェック</param>
        /// <param name="other_voice_color">その他ボイスカラー</param>
        /// <param name="voice_color">ボイスカラー</param>
        /// <param name="font_name">フォント名</param>
        /// <param name="version">バージョン</param>
        /// <param name="today_date">日付</param>
        /// <param name="convert_comment">コメント文字を変換する文字</param>
        /// <param name="check_star">☆追加チェック</param>
        /// <param name="excel_visible">エクセル可視チェック</param>
        public ExcelData(int zoom, int font_size, string comment, string title, bool bold, 
            Color other_voice_color, Color voice_color, string font_name,
            string version, string today_date, bool convert_comment, bool check_star, bool excel_visible)
        {
            this.zoom = zoom;
            this.font_size = font_size;
            this.comment = comment;
            this.title = title;
            this.bold = bold;
            this.other_voice_color = other_voice_color;
            this.voice_color = voice_color;
            this.font_name = font_name;
            this.version = version;
            this.today_date = today_date;
            this.convert_comment = convert_comment;
            this.check_star = check_star;
            this.excel_visible = excel_visible;
        }
    };
}

//汎用操作
namespace util
{
    static class Util
    {
        public static bool CreateDirectory(string directory)
        {
            try
            {
                if (Directory.Exists(directory))
                {
                    Directory.Delete(directory, true);
                }
                Directory.CreateDirectory(directory);

            }
            catch (Exception e)
            {
                Debug.WriteLine("/******************************************************************************************************************");
                Debug.WriteLine(e + "エラー発生っす! : CreateDirectory");
                Debug.WriteLine("******************************************************************************************************************/");
                return false;
            }
            return true;
        }
        public static string[] NowProcess()
        {
            List<string> s = new List<string>();
            System.Diagnostics.Process[] ps = System.Diagnostics.Process.GetProcesses();

            foreach (var p in ps)
            {
                try
                {
                    s.Add(p.ProcessName);
                    s.Add(p.MainModule.FileName);
                }
                catch (Exception e)
                {
                    Debug.WriteLine(e + "エラー : NowProcess");
                }
            }

            return s.ToArray();
        }
        public static string TodayUpdateString()
        {
            return DateTime.Today.ToShortDateString() + "\n";
        }
        public static bool FileCheck(string path, string extension)
        {
            if (File.Exists(path))
            {
                if (System.IO.Path.GetExtension(path) == extension)
                {
                    return true;
                }
            }
            return false;
        }
        public static void FileCopy(string src_file, string dst_file)
        {
            System.IO.File.Copy(src_file, dst_file, true);
        }
        //プロセスチェック
        //http://cloudyheaven.blog130.fc2.com/blog-entry-100.html
        public static bool ProcessCheck(string path)
        {
            try
            {
                //File.Open(path, FileMode.Open);
                var fs = new FileStream(path, FileMode.Open, FileAccess.Read);
                fs.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine("ファイルが別のプロセスで使用中のため、オープンできません。: " + e);
                return false;
            }
            return true;
        }
        public static bool ProcessCheck2(string path)
        {
            if (Process.GetProcessesByName(path).Length != 0)
            {
                //Process.Start(path);
                return false;
            }
            return true;
        }

        //http://dobon.net/vb/dotnet/string/detectcode.html
        public static System.Text.Encoding GetCode(byte[] bytes)
        {
            const byte bEscape = 0x1B;
            const byte bAt = 0x40;
            const byte bDollar = 0x24;
            const byte bAnd = 0x26;
            const byte bOpen = 0x28;    //'('
            const byte bB = 0x42;
            const byte bD = 0x44;
            const byte bJ = 0x4A;
            const byte bI = 0x49;

            int len = bytes.Length;
            byte b1, b2, b3, b4;

            //Encode::is_utf8 は無視

            bool isBinary = false;
            for (int i = 0; i < len; i++)
            {
                b1 = bytes[i];
                if (b1 <= 0x06 || b1 == 0x7F || b1 == 0xFF)
                {
                    //'binary'
                    isBinary = true;
                    if (b1 == 0x00 && i < len - 1 && bytes[i + 1] <= 0x7F)
                    {
                        //smells like raw unicode
                        return System.Text.Encoding.Unicode;
                    }
                }
            }
            if (isBinary)
            {
                return null;
            }

            //not Japanese
            bool notJapanese = true;
            for (int i = 0; i < len; i++)
            {
                b1 = bytes[i];
                if (b1 == bEscape || 0x80 <= b1)
                {
                    notJapanese = false;
                    break;
                }
            }
            if (notJapanese)
            {
                return System.Text.Encoding.ASCII;
            }

            for (int i = 0; i < len - 2; i++)
            {
                b1 = bytes[i];
                b2 = bytes[i + 1];
                b3 = bytes[i + 2];

                if (b1 == bEscape)
                {
                    if (b2 == bDollar && b3 == bAt)
                    {
                        //JIS_0208 1978
                        //JIS
                        return System.Text.Encoding.GetEncoding(50220);
                    }
                    else if (b2 == bDollar && b3 == bB)
                    {
                        //JIS_0208 1983
                        //JIS
                        return System.Text.Encoding.GetEncoding(50220);
                    }
                    else if (b2 == bOpen && (b3 == bB || b3 == bJ))
                    {
                        //JIS_ASC
                        //JIS
                        return System.Text.Encoding.GetEncoding(50220);
                    }
                    else if (b2 == bOpen && b3 == bI)
                    {
                        //JIS_KANA
                        //JIS
                        return System.Text.Encoding.GetEncoding(50220);
                    }
                    if (i < len - 3)
                    {
                        b4 = bytes[i + 3];
                        if (b2 == bDollar && b3 == bOpen && b4 == bD)
                        {
                            //JIS_0212
                            //JIS
                            return System.Text.Encoding.GetEncoding(50220);
                        }
                        if (i < len - 5 &&
                            b2 == bAnd && b3 == bAt && b4 == bEscape &&
                            bytes[i + 4] == bDollar && bytes[i + 5] == bB)
                        {
                            //JIS_0208 1990
                            //JIS
                            return System.Text.Encoding.GetEncoding(50220);
                        }
                    }
                }
            }

            //should be euc|sjis|utf8
            //use of (?:) by Hiroki Ohzaki <ohzaki@iod.ricoh.co.jp>
            int sjis = 0;
            int euc = 0;
            int utf8 = 0;
            for (int i = 0; i < len - 1; i++)
            {
                b1 = bytes[i];
                b2 = bytes[i + 1];
                if (((0x81 <= b1 && b1 <= 0x9F) || (0xE0 <= b1 && b1 <= 0xFC)) &&
                    ((0x40 <= b2 && b2 <= 0x7E) || (0x80 <= b2 && b2 <= 0xFC)))
                {
                    //SJIS_C
                    sjis += 2;
                    i++;
                }
            }
            for (int i = 0; i < len - 1; i++)
            {
                b1 = bytes[i];
                b2 = bytes[i + 1];
                if (((0xA1 <= b1 && b1 <= 0xFE) && (0xA1 <= b2 && b2 <= 0xFE)) ||
                    (b1 == 0x8E && (0xA1 <= b2 && b2 <= 0xDF)))
                {
                    //EUC_C
                    //EUC_KANA
                    euc += 2;
                    i++;
                }
                else if (i < len - 2)
                {
                    b3 = bytes[i + 2];
                    if (b1 == 0x8F && (0xA1 <= b2 && b2 <= 0xFE) &&
                        (0xA1 <= b3 && b3 <= 0xFE))
                    {
                        //EUC_0212
                        euc += 3;
                        i += 2;
                    }
                }
            }
            for (int i = 0; i < len - 1; i++)
            {
                b1 = bytes[i];
                b2 = bytes[i + 1];
                if ((0xC0 <= b1 && b1 <= 0xDF) && (0x80 <= b2 && b2 <= 0xBF))
                {
                    //UTF8
                    utf8 += 2;
                    i++;
                }
                else if (i < len - 2)
                {
                    b3 = bytes[i + 2];
                    if ((0xE0 <= b1 && b1 <= 0xEF) && (0x80 <= b2 && b2 <= 0xBF) &&
                        (0x80 <= b3 && b3 <= 0xBF))
                    {
                        //UTF8
                        utf8 += 3;
                        i += 2;
                    }
                }
            }
            //M. Takahashi's suggestion
            //utf8 += utf8 / 2;

            System.Diagnostics.Debug.WriteLine(
                string.Format("sjis = {0}, euc = {1}, utf8 = {2}", sjis, euc, utf8));
            if (euc > sjis && euc > utf8)
            {
                //EUC
                return System.Text.Encoding.GetEncoding(51932);
            }
            else if (sjis > euc && sjis > utf8)
            {
                //SJIS
                return System.Text.Encoding.GetEncoding(932);
            }
            else if (utf8 > euc && utf8 > sjis)
            {
                //UTF8
                return System.Text.Encoding.UTF8;
            }

            return null;
        }
    }
}
