using System.IO;
using System.Xml.Serialization;
using System.Collections;

namespace LibrettoCreateTool
{

    public partial class MyXmlSettings
    {
        #region プロパティ
        /// <summary>
        /// プロパティ
        /// </summary>
        public string SrcDir { get { return _src_dir; } set { _src_dir = value; } }
        /// <summary>
        /// プロパティ
        /// </summary>
        public string DstDir { get { return _dst_dir; } set { _dst_dir = value; } }
        /// <summary>
        /// プロパティ
        /// </summary>
        public string[][] Cells { get { return _cells; } set { _cells = value; } }
        /// <summary>
        /// プロパティ
        /// </summary>
        public string Comment { get { return _comment; } set { _comment = value; } }
        /// <summary>
        /// プロパティ
        /// </summary>
        public string Font { get { return _font; } set { _font = value; } }
        /// <summary>
        /// プロパティ
        /// </summary>
        public string Title { get { return _title; } set { _title = value; } }
        /// <summary>
        /// プロパティ
        /// </summary>
        public string LogoText { get { return _logo_text; } set { _logo_text = value; } }
        /// <summary>
        /// プロパティ
        /// </summary>
        public int LogoPos { get { return _logo_pos; } set { _logo_pos = value; } }
        /// <summary>
        /// プロパティ
        /// </summary>
        public int PagePos { get { return _page_pos; } set { _page_pos = value; } }
        /// <summary>
        /// プロパティ
        /// </summary>
        public int FilePos { get { return _file_pos; } set { _file_pos = value; } }
        /// <summary>
        /// プロパティ
        /// </summary>
        public int KeyPos { get { return _key_pos; } set { _key_pos = value; } }
        /// <summary>
        /// プロパティ
        /// </summary>
        public string VoiceTop { get { return _voice_top; } set { _voice_top = value; } }
        /// <summary>
        /// プロパティ
        /// </summary>
        public string VoiceMiddle { get { return _voice_middle; } set { _voice_middle = value; } }
        /// <summary>
        /// プロパティ
        /// </summary>
        public string VoiceBottom { get { return _voice_bottom; } set { _voice_bottom = value; } }
        /// <summary>
        /// プロパティ
        /// </summary>
        public bool VoiceBold { get { return _voice_bold; } set { _voice_bold = value; } }
        /// <summary>
        /// プロパティ
        /// </summary>
        public int VoiceOtherColor { get { return _voice_other_color; } set { _voice_other_color = value; } }
        /// <summary>
        /// プロパティ
        /// </summary>
        public int VoiceColor { get { return _voice_color; } set { _voice_color = value; } }
        /// <summary>
        /// プロパティ
        /// </summary>
        public string Version { get { return _version; } set { _version = value; } }
        /// <summary>
        /// プロパティ
        /// </summary>
        public bool CommentConvert { get { return _comment_convert; } set { _comment_convert = value; } }
        /// <summary>
        /// プロパティ
        /// </summary>
        public bool CheckStar { get { return _check_add_start; } set { _check_add_start = value; } }
        /// <summary>
        /// プロパティ
        /// </summary>
        public bool Visible { get { return _excel_visible; } set { _excel_visible = value; } }
        /// <summary>
        /// チェックボックス
        /// </summary>
        public bool CheckBox1{ get { return _check_box1;} set { _check_box1 = value; }}
        public bool CheckBox2{ get { return _check_box2;} set { _check_box2 = value; }}
        public bool CheckBox3{ get { return _check_box3;} set { _check_box3 = value; }}
        public bool CheckBox4{ get { return _check_box4;} set { _check_box4 = value; }}
        public bool CheckBox5{ get { return _check_box5;} set { _check_box5 = value; }}

        //リクエストトークンはXMLに書き込まない
        //[System.Xml.Serialization.XmlIgnoreAttribute]
        #endregion
    }
    /// <summary>
    /// xml設定ファイル作成用クラス
    /// </summary>
    public partial class MyXmlSettings
    {
        //http://kuroeveryday.blogspot.jp/2013/06/csharpxml.html
        //http://www.kanazawa-net.ne.jp/~pmansato/net/net_tech_serialize.htm

        #region メンバ変数
        string _src_dir;
        string _dst_dir;
        string[][] _cells;

        string _comment;
        string _font;
        string _title;
        string _logo_text;

        int _logo_pos;
        int _page_pos;
        int _file_pos;
        int _key_pos;

        string _voice_top;
        string _voice_middle;
        string _voice_bottom;

        bool _voice_bold;
        int _voice_other_color;
        int _voice_color;
        string _version;
        bool _comment_convert;
        bool _check_add_start;
        bool _excel_visible;

        bool _check_box1;
        bool _check_box2;
        bool _check_box3;
        bool _check_box4;
        bool _check_box5;

        #endregion
        /// <summary>
        /// コンストラクタ(空)
        /// </summary>
        public MyXmlSettings() { }
        /// <summary>
        /// 値設定
        /// </summary>
        /// <param name="xml">xmlデータ</param>
        private void SetMyXml(MyXmlSettings xml)
        {
            this.SrcDir = xml.SrcDir;
            this.DstDir = xml.DstDir;
            this.Cells = xml.Cells;
            this.Comment = xml.Comment;
            this.Font = xml.Font;
            this.Title = xml.Title;
            this.LogoText = xml.LogoText;

            this.LogoPos = xml.LogoPos;
            this.PagePos = xml.PagePos;
            this.FilePos = xml.FilePos;
            this.KeyPos = xml.KeyPos;

            this.VoiceTop = xml.VoiceTop;
            this.VoiceMiddle = xml.VoiceMiddle;
            this.VoiceBottom = xml.VoiceBottom;

            this.VoiceBold = xml.VoiceBold;
            this.VoiceOtherColor = xml.VoiceOtherColor;
            this.VoiceColor = xml.VoiceColor;
            this.Version = xml.Version;

            this.CommentConvert = xml.CommentConvert;
            this.CheckStar = xml.CheckStar;
            this.Visible = xml.Visible;

            this.CheckBox1 = xml.CheckBox1;
            this.CheckBox2 = xml.CheckBox2;
            this.CheckBox3 = xml.CheckBox3;
            this.CheckBox4 = xml.CheckBox4;
            this.CheckBox5 = xml.CheckBox5;
        }
        /// <summary>
        /// xml書き込み
        /// </summary>
        /// <param name="path_xml">xmlのパス</param>
        /// <returns></returns>
        public bool WriteXml(string path_xml)
        {
            var serializer = new XmlSerializer(typeof(MyXmlSettings));
            using (var fs = new FileStream(path_xml, FileMode.Create))
            {
                serializer.Serialize(fs, this);
            }
            return true;
        }
        /// <summary>
        /// xml読み込み
        /// </summary>
        /// <param name="path_xml">xmlのパス</param>
        /// <returns></returns>
        public bool ReadXml(string path_xml)
        {
            if (!FileCheck(path_xml)) return false;

            var serializer = new XmlSerializer(typeof(MyXmlSettings));
            using (var fs = new FileStream(path_xml, FileMode.Open))
            {
                var data = (MyXmlSettings)serializer.Deserialize(fs);

                SetMyXml(data);
            }
            return true;
        }
        /// <summary>
        /// xml用ファイルチェック
        /// </summary>
        /// <param name="path">ファイルパス</param>
        /// <returns></returns>
        public bool FileCheck(string path)
        {
            if (util.Util.FileCheck(path, ".xml")) return true;
            return false;
        }
    }
}
