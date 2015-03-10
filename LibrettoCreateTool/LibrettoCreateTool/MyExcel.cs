using System;
using System.IO;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;       // 参照の追加とこれを記述
using System.Diagnostics;
using System.Drawing;

namespace LibrettoCreateTool
{
    public partial class MyExcel
    {
        #region 印刷用enum
        /// <summary>
        /// ヘッダーフッター用ポジション
        /// </summary>
        public enum PRINT_POSITION
        {
            /// <summary>
            /// 左
            /// </summary>
            LEFT,
            /// <summary>
            /// 真ん中
            /// </summary>
            CENTER,
            /// <summary>
            /// 右
            /// </summary>
            RIGHT,
        };
        /// <summary>
        /// ヘッダーフッター用
        /// </summary>
        public enum PRINT_HEAD_FOOT
        {
            /// <summary>
            /// ヘッダー
            /// </summary>
            HEADER,
            /// <summary>
            /// フッター
            /// </summary>
            FOOTER,
        };
        /// <summary>
        /// 印刷用 プリントサイズ
        /// </summary>
        public enum PRINT_SIZE
        {
            /// <summary>
            /// 印刷のプリントサイズ
            /// </summary>
            A3 = Excel.XlPaperSize.xlPaperA3,
            /// <summary>
            /// 印刷のプリントサイズ
            /// </summary>
            A4 = Excel.XlPaperSize.xlPaperA4,
            /// <summary>
            /// 印刷のプリントサイズ
            /// </summary>
            A5 = Excel.XlPaperSize.xlPaperA5,
            /// <summary>
            /// 印刷のプリントサイズ
            /// </summary>
            B4 = Excel.XlPaperSize.xlPaperB4,
            /// <summary>
            /// 印刷のプリントサイズ
            /// </summary>
            B5 = Excel.XlPaperSize.xlPaperB5,
        };
        #endregion
        #region セル用enum
        /// <summary>
        /// セル行列用
        /// </summary>
        public enum EXCEL_MATRIX
        {
            /// <summary>
            /// 行
            /// </summary>
            X,
            /// <summary>
            /// 列
            /// </summary>
            Y,
        };
        /// <summary>
        /// セル調整
        /// Horizon (水平)
        /// </summary>
        public enum CELL_H_DIR
        {            
            /// <summary>
            /// 左詰め
            /// </summary>
            LEFT = Excel.XlHAlign.xlHAlignLeft,
            /// <summary>
            /// 右詰め
            /// </summary>
            RIGHT = Excel.XlHAlign.xlHAlignRight,
            /// <summary>
            /// 中央
            /// </summary>
            CENTER = Excel.XlHAlign.xlHAlignCenter,
        };
        /// <summary>
        /// セル調整
        /// Vertical (垂直)
        /// </summary>
        public enum CELL_V_DIR
        {
            /// <summary>
            /// 上詰め
            /// </summary>
            TOP = Excel.XlVAlign.xlVAlignTop,            
            /// <summary>
            /// 下詰め
            /// </summary>
            BOTTOM = Excel.XlVAlign.xlVAlignBottom,
            /// <summary>
            /// 中央
            /// </summary>
            CENTER = Excel.XlVAlign.xlVAlignCenter,
        };
        #endregion
        #region デザイン用enum
        /// <summary>
        /// 罫線描画
        /// </summary>
        public enum BORDER_RANGE
        {
            /// <summary>
            /// 外枠の下
            /// </summary>
            EDGE_BOTTOM = Excel.XlBordersIndex.xlEdgeBottom,
            /// <summary>
            /// 外枠の左
            /// </summary>
            EDGE_LEFT = Excel.XlBordersIndex.xlEdgeLeft,
            /// <summary>
            /// 外枠の右
            /// </summary>
            EDGE_RIGHT = Excel.XlBordersIndex.xlEdgeRight,
            /// <summary>
            /// 外枠の上
            /// </summary>
            EDGE_TOP = Excel.XlBordersIndex.xlEdgeTop,
            /// <summary>
            /// 内枠の水平方向
            /// </summary>
            INSIDE_HORIZONTAL = Excel.XlBordersIndex.xlInsideHorizontal,
            /// <summary>
            /// 内枠の垂直方向
            /// </summary>
            INSIDE_VERTICAL = Excel.XlBordersIndex.xlInsideVertical,
        };
        #endregion
    }
    public partial class MyExcel
    {
        #region メンバ変数
        private Excel.Application excel_ = null;
        private Excel.Workbook book_ = null;
        private Excel.Worksheet sheet_ = null;
        private object[,] cells_;               //Excelファイルへのアクセス数を減らすための仮想セル ※セル１つ１つに値をセットしていくと、非常に遅くなるため。
        private int cell_max_x_;
        private int cell_max_y_;
        private int cell_start_x_;
        private int cell_start_y_;
        #endregion
    }
    public partial class MyExcel
    {
        #region 拡張用関数 (現在使用していません)
        /// <summary>
        /// /// ワークシートコピー
        /// </summary>
        /// <param name="work_index">コピー元</param>
        /// <param name="copy_index">コピー先</param>
        /// <param name="sheet_name">シート名</param>
        private void CopyWorksheet(int work_index, int copy_index, string sheet_name)
        {
            if (sheet_ != null)
            {
                Excel.Worksheet copy_work = (Excel.Worksheet)book_.Worksheets[work_index];
                copy_work.Name = sheet_name;
                copy_work.Copy(System.Reflection.Missing.Value, book_.Worksheets[copy_index]);
            }
        }
        /// <summary>
        /// ワークシートコピー 最後に追加
        /// </summary>
        /// <param name="work_index">コピー元</param>
        /// <param name="sheet_name">シート名</param>
        private void CopyWorksheet(int work_index, string sheet_name)
        {
            if (sheet_ != null)
            {
                Excel.Worksheet copy_work = (Excel.Worksheet)book_.Worksheets[work_index];
                copy_work.Name = sheet_name;
                copy_work.Copy(System.Reflection.Missing.Value, book_.Worksheets[GetSheetSize()]);
            }
        }
        /// <summary>
        /// ワークシートの削除
        /// </summary>
        /// <param name="work_index">削除インデックス</param>
        private void DeleteWorksheet(int work_index)
        {
            (excel_.ActiveWorkbook.Sheets[work_index]).Delete();
        }
        /// <summary>
        /// ワークシートの削除 1枚残さないと例外が発生するので、１枚残ります
        /// </summary>
        private void AllDeleteWorksheet()
        {
            while (book_.Worksheets.Count > 1)
            {
                (excel_.ActiveWorkbook.Sheets[book_.Worksheets.Count]).Delete();
            }
        }
        #endregion
    }
    /// <summary>
    /// エクセル用クラス
    /// </summary>
    public partial class MyExcel
    {
        //readme
        //http://www.red.oit-net.jp/tatsuya/vb/Excel.htm
        //http://www.eurus.dti.ne.jp/~yoneyama/Excel/vba/vba_font.html
        //http://www.excellenceweb.net/vba/object/range_member/font/fontstyle.html
        #region 関数群
        #region 確定処理
        /// <summary>
        /// エクセルアプリケーションの生成
        /// </summary>
        public MyExcel()
        {
            excel_ = new Excel.Application();
        }
        /// <summary>
        /// 解放処理を行います
        /// </summary>
        public void End()
        {
            if (sheet_ != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet_);
                sheet_ = null;
            }
            if (book_ != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(book_);
                book_ = null;
            }
            if (excel_ != null)
            {
                excel_.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel_);
                excel_ = null;
                //確実にオブジェクト削除 //※完全終了する場合のみ
                System.GC.Collect();
            }
        }
        /// <summary>
        /// Excelを可視状態にします
        /// </summary>
        public void Visible()
        {
            excel_.Visible = true;
        }
        /// <summary>
        /// Excelの警告ダイアログを無視します
        /// </summary>
        public void Alerts()
        {
            excel_.DisplayAlerts = false;
        }
        #endregion
        #region ブック設定
        /// <summary>
        /// ブックを生成します
        /// </summary>
        public void CreateBook()
        {
            book_ = (Excel.Workbook)(excel_.Workbooks.Add(Type.Missing));
            sheet_ = book_.ActiveSheet;
        }
        /// <summary>
        /// ブックを生成します シート枚数の設定 
        /// </summary>
        /// <param name="sheet_max"></param>
        public void CreateBook(int sheet_max)
        {
            book_ = (Excel.Workbook)(excel_.Workbooks.Add(sheet_max));
            sheet_ = book_.ActiveSheet;
        }
        /// <summary>
        /// ブックを閉じます
        /// </summary>
        public void Close()
        {
            if (sheet_ != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet_);
                sheet_ = null;
            }
            book_.Close(false);
            if (book_ != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(book_);
                book_ = null;
            }
        }
        /// <summary>
        /// xls形式で保存を行います
        /// </summary>
        /// <param name="file_name"></param>
        public void SaveXls(string file_name)
        {
            if (File.Exists(file_name))
            {
                File.Delete(file_name);
            }
            book_.SaveAs(file_name, Excel.XlFileFormat.xlExcel7, Type.Missing, Type.Missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        }
        /// <summary>
        /// xlsx形式で保存を行います
        /// </summary>
        /// <param name="file_name"></param>
        public void SaveXlsx(string file_name)
        {
            if (File.Exists(file_name))
            {
                File.Delete(file_name);
            }
            book_.SaveAs(file_name, Type.Missing, Type.Missing, Type.Missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        }
        /// <summary>
        /// ブックを開きます
        /// </summary>
        /// <param name="file_name"></param>
        public void Open(string file_name)
        {
            book_ = (Excel.Workbook)(excel_.Workbooks.Open(file_name,
                Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing));
        }
        #endregion
        #region シート設定
        /// <summary>
        /// シート名からシート番号を探します。
        /// 失敗 -1
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        private int getSheetIndex(string sheetName)
        {
            int i = 0;
            foreach (Excel.Worksheet sh in book_.Sheets)
            {
                if (sheetName == sh.Name)
                {
                    return i + 1;
                }
                i += 1;
            }
            return -1;
        }
        /// <summary>
        /// 現在のアクティブシートを削除します。
        /// </summary>
        public void SheetDelete()
        {
            sheet_.Delete();
        }
        /// <summary>
        /// 指定されたインデックスのシートを削除します
        /// </summary>
        /// <param name="index">削除するシートのインデックス</param>
        public void SheetDelete(int index)
        {
            (book_.Sheets[index]).Delete();
        }
        /// <summary>
        /// 仮想セルに値をセットする
        /// </summary>
        public void SetRange()
        {
            sheet_.Range[RangeTop_(), RangeButton_()].Value2 = cells_;
        }
        /// <summary>
        /// 字詰め調整 LEFT RIGHT CENTER
        /// </summary>
        /// <param name="dir"></param>
        public void CellAlignment(CELL_H_DIR dir)
        {
            sheet_.get_Range(RangeTop_(), RangeButton_()).HorizontalAlignment = dir;
        }
        /// <summary>
        /// 字詰め調整 LEFT RIGHT CENTER
        /// </summary>
        /// <param name="dir"></param>
        /// <param name="mat">EXCEL_MATRIXでの1行 or 1列指定</param>
        /// <param name="index">その行 or 列の場所指定</param>
        public void CellAlignment(CELL_H_DIR dir, EXCEL_MATRIX mat, int index)
        {
            sheet_.get_Range(RangeTop_(mat, index), RangeButton_(mat, index)).HorizontalAlignment = dir;
        }
        /// <summary>
        /// 字詰め調整 TOP BOTTOM CENTER
        /// </summary>
        /// <param name="dir"></param>
        public void CellAlignment(CELL_V_DIR dir)
        {
            sheet_.get_Range(RangeTop_(), RangeButton_()).VerticalAlignment = dir;
        }
        /// <summary>
        /// 字詰め調整 LEFT RIGHT CENTER
        /// </summary>
        /// <param name="dir"></param>
        /// <param name="mat">EXCEL_MATRIXでの1行 or 1列指定</param>
        /// <param name="index">その行 or 列の場所指定</param>
        public void CellAlignment(CELL_V_DIR dir, EXCEL_MATRIX mat, int index)
        {
            sheet_.get_Range(RangeTop_(mat, index), RangeButton_(mat, index)).VerticalAlignment = dir;
        }
        /// <summary>
        /// セル内折り返し機能
        /// </summary>
        public void CellWrapMode()
        {
            sheet_.get_Range(RangeTop_(), RangeButton_()).WrapText = true;
        }
        /// <summary>
        /// セル結合を行います
        /// </summary>
        public void CellMerge()
        {
            sheet_.get_Range(RangeTop_(), RangeButton_()).MergeCells = true;
        }
        /// <summary>
        /// シート名変更
        /// </summary>
        /// <param name="_name"></param>
        public void SetSheetName(string _name)
        {
            sheet_.Name = _name;
        }
        /// <summary>
        /// シートを最後のシートに設定します
        /// </summary>
        public void SetLastSheet()
        {
            sheet_ = (Excel.Worksheet)book_.Worksheets[book_.Worksheets.Count];
        }
        /// <summary>
        /// 現在のブックの最後のシートを取得します
        /// </summary>
        /// <param name="sheet_name"></param>
        public void SetLastSheet(string sheet_name)
        {
            sheet_ = (Excel.Worksheet)book_.Worksheets[book_.Worksheets.Count];
            sheet_.Name = sheet_name;
        }
        /// <summary>
        /// シートをセットします
        /// </summary>
        /// <param name="index">シートインデックス指定</param>
        public void SetSheet(int index)
        {
            sheet_ = (Excel.Worksheet)book_.Worksheets[index];
        }
        /// <summary>
        /// 新規のシートを作成します
        /// </summary>
        public void CreateNewSheet(string _sheet_name)
        {
            sheet_ = (Excel.Worksheet)book_.Worksheets.Add(Type.Missing, sheet_, 1, Type.Missing);
            sheet_.Name = _sheet_name;
        }
        /// <summary>
        /// シートを作成します
        /// </summary>
        /// <param name="index">シート数</param>
        public void CreateSheet(int index)
        {
            sheet_ = new Excel.Worksheet();
            if (index > book_.Sheets.Count || index == 0) index = 1;
            sheet_ = (Excel.Worksheet)book_.Sheets[index];
        }
        /// <summary>
        /// シートの最大数を取得します
        /// </summary>
        /// <returns></returns>
        public int GetSheetSize()
        {
            return book_.Sheets.Count;
        }
        #endregion
        #region セル設定
        /// <summary>
        /// 仮想セルにデータを設定します
        /// </summary>
        /// <param name="y">yのインデックス</param>
        /// <param name="x">xのインデックス</param>
        /// <param name="data"></param>
        public void SetCell(int y, int x, object data)
        {
            cells_[y, x] = data;
        }
        /// <summary>
        /// 仮想セルを生成します
        /// </summary>
        /// <param name="start_y"></param>
        /// <param name="start_x"></param>
        /// <param name="end_y"></param>
        /// <param name="end_x"></param>
        public void CreateCell(int start_y, int start_x, int end_y, int end_x)
        {
            cell_start_y_ = start_y;
            cell_start_x_ = start_x;
            cell_max_y_ = end_y;
            cell_max_x_ = end_x;
            cells_ = new object[end_y, end_x];
        }
        /// <summary>
        /// 仮想セルを生成(渡します)
        /// </summary>
        /// <param name="start_y"></param>
        /// <param name="start_x"></param>
        /// <param name="cell"></param>
        public void CreateCell(int start_y, int start_x, object[,] cell)
        {
            cell_start_y_ = start_y;
            cell_start_x_ = start_x;
            cell_max_y_ = cell.GetLength(0);
            cell_max_x_ = cell.GetLength(1);
            cells_ = cell;
        }
        /// <summary>
        /// Xの最大値
        /// </summary>
        /// <returns></returns>
        public int EndX()
        {
            return cell_max_x_;
        }
        /// <summary>
        /// Yの最大値
        /// </summary>
        /// <returns></returns>
        public int EndY()
        {
            return cell_max_y_;
        }
        /// <summary>
        /// Xのスタート値
        /// </summary>
        /// <returns></returns>
        public int StartX()
        {
            return cell_start_x_;
        }
        /// <summary>
        /// Yのスタート値
        /// </summary>
        /// <returns></returns>
        public int StartY()
        {
            return cell_start_y_;
        }
        /// <summary>
        /// 回転
        /// </summary>
        public void ChangeXY()
        {
            int max = EndX() * EndY();
            if (max == 0)
            {
                Console.WriteLine("行列じゃないと設定できないです。");
                return;
            }
            object[] temp = new object[max];
            object[] temp2 = new object[max];
            for (int i = 0; i < EndY(); ++i)
            {
                for (int j = 0; j < EndX(); ++j)
                {
                    temp2[i * EndX() + j] = temp[i * EndX() + j] = cells_[i, j];
                }
            }
            int cnt = 0;
            for (int i = EndY() - 1; i >= 0; --i)
            {
                for (int j = 0; j < EndX(); ++j)
                {
                    temp[i + j * EndY()] = temp2[cnt++];
                }
            }
            CreateCell(StartX(), StartY(), EndX(), EndY());
            for (uint i = 0; i < EndY(); ++i)
            {
                for (uint j = 0; j < EndX(); ++j)
                {
                    cells_[i, j] = temp[i * EndX() + j];
                }
            }
        }
        /// <summary>
        /// 回転
        /// </summary>
        public void ChangeXY2()
        {
            int max = EndX() * EndY();
            if (max == 0)
            {
                Console.WriteLine("行列じゃないと設定できないです。");
                return;
            }
            object[] temp = new object[max];
            object[] temp2 = new object[max];
            for (int i = 0; i < EndY(); ++i)
            {
                for (int j = 0; j < EndX(); ++j)
                {
                    temp2[i * EndX() + j] = temp[i * EndX() + j] = cells_[i, j];
                }
            }
            int cnt = 0;
            for (int i = 0; i < EndY(); ++i)
            {
                for (int j = 0; j < EndX(); ++j)
                {
                    temp[i + j * EndY()] = temp2[cnt++];
                }
            }
            CreateCell(StartX(), StartY(), EndX(), EndY());
            for (uint i = 0; i < EndY(); ++i)
            {
                for (uint j = 0; j < EndX(); ++j)
                {
                    cells_[i, j] = temp[i * EndX() + j];
                }
            }
        }
        /// <summary>
        /// 値をセルに変更
        /// </summary>
        /// <param name="index">セルインデックス</param>
        /// <returns></returns>
        private string GetColumnName_(int index)
        {
            string str = "";
            do
            {
                str = Convert.ToChar(index % 26 + 0x41) + str;
            } while ((index = index / 26 - 1) != -1);

            return str;
        }
        /// <summary>
        /// 列変更
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        private string RangeTopX_(int index)
        {
            return GetColumnName_(index) + (cell_start_y_ + 1).ToString();
        }
        /// <summary>
        /// 行変更
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        private string RangeTopY_(int index)
        {
            return GetColumnName_(cell_start_x_) + (index + 1).ToString();
        }
        /// <summary>
        /// セル範囲の先頭文字列生成
        /// </summary>
        /// <returns></returns>
        private string RangeTop_()
        {
            return GetColumnName_(cell_start_x_) + (cell_start_y_ + 1).ToString();  //セルの始まりは(0, 0)ではなく(1, 1)のため、+ 1
        }
        /// <summary>
        /// セル範囲の末尾文字列生成
        /// </summary>
        /// <returns></returns>
        private string RangeButton_()
        {
            //スタート位置分ずらす
            int max_x = cell_max_x_ + cell_start_x_;
            int max_y = cell_max_y_ + cell_start_y_;
            return GetColumnName_(max_x - 1) + (max_y).ToString();
        }
        /// <summary>
        /// 指定されたインデックスのセル文字列生成
        /// </summary>
        /// <param name="y"></param>
        /// <param name="x"></param>
        /// <returns></returns>
        private string RangeCell_(int y, int x)
        {
            var s = GetColumnName_(x) + (y + 1).ToString();
            return s;
        }
        #region Rangeデリゲート使用
        delegate string RangeD1(int index);
        delegate string RangeD2(int index, int top_bottom);
        /// <summary>
        /// 範囲指定
        /// </summary>
        /// <param name="mat">EXCEL_MATRIXでの1行 or 1列指定</param>
        /// <param name="index">その行 or 列の場所指定</param>
        public void SetRange(EXCEL_MATRIX mat, int index)
        {
            //sheet_.get_Range(RangeTop_(mat, index), RangeButton_(mat, index)).Value2 = cells_;
            Excel.Range range = sheet_.get_Range(RangeTop_(mat, index), RangeButton_(mat, index));
            range.Value2 = cells_;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
        }
        private string RangeTop_(EXCEL_MATRIX mat, int index)
        {
            RangeD1[] rd = 
            {
                RangeTopRow_,
                RangeTopColumn_,
            };
            return rd[(int)mat](index);
        }
        private string RangeTop_(EXCEL_MATRIX mat, int range, int top)
        {
            RangeD2[] rd = 
            {
                RangeTopRow_,
                RangeTopColumn_,
            };
            return rd[(int)mat](range, top);
        }
        private string RangeButton_(EXCEL_MATRIX mat, int range, int bottom)
        {
            RangeD2[] rd = 
            {
                RangeBottomRow_,
                RangeBottomColumn_,                
            };
            return rd[(int)mat](range, bottom);
        }
        private string RangeTopColumn_(int range, int top)
        {
            return GetColumnName_(top) + (range + 1).ToString();
        }
        private string RangeBottomColumn_(int index, int bottom)
        {
            return GetColumnName_(index) + (bottom).ToString();
        }
        private string RangeTopRow_(int range, int top)
        {
            return GetColumnName_(top) + (range + 1).ToString();
        }
        private string RangeBottomRow_(int range, int bottom)
        {
            return GetColumnName_(bottom) + (range + 1).ToString();
        }
        private string RangeButton_(EXCEL_MATRIX mat, int index)
        {
            RangeD1[] rd =
            {
                RangeButtonRow_,
                RangeButtonColumn_,
            };
            return rd[(int)mat](index);
        }
        private string RangeTopRow_(int index)
        {
            return GetColumnName_(cell_start_x_) + (index + 1).ToString();
        }
        private string RangeButtonRow_(int index)
        {
            //スタート位置分ずらす
            int max_x = cell_max_x_ + cell_start_x_;
            int max_y = cell_max_y_ + cell_start_y_;
            return GetColumnName_(max_x - 1) + (index + 1).ToString();
        }
        private string RangeTopColumn_(int index)
        {
            return GetColumnName_(index) + (cell_start_y_ + 1).ToString();
        }
        private string RangeButtonColumn_(int index)
        {
            int max_x = cell_max_x_ + cell_start_x_;
            int max_y = cell_max_y_ + cell_start_y_;
            return GetColumnName_(index) + (max_y).ToString();
        }
        #endregion
        #endregion
        #region デザイン設定
        #region フォントサイズ
        /// <summary>
        /// フォントサイズ
        /// </summary>
        /// <param name="font_size"></param>
        public void SetFontSize(int font_size)
        {
            sheet_.get_Range(RangeTop_(), RangeButton_()).Font.Size = font_size;
        }
        /// <summary>
        /// フォントサイズ
        /// </summary>
        /// <param name="mat">EXCEL_MATRIXでの1行 or 1列指定</param>
        /// <param name="index">その行 or 列の場所指定</param>
        /// <param name="font_size"></param>
        private void SetFontSize_(EXCEL_MATRIX mat, int index, int font_size)
        {
            sheet_.get_Range(RangeTop_(mat, index), RangeButton_(mat, index)).Font.Size = font_size;
        }
        #endregion
        #region ボーダー設定
        /// <summary>
        /// ボーダー色設定
        /// </summary>
        /// <param name="color">カラー指定</param>
        public void SetBorderColor(Color color)
        {
            sheet_.get_Range(RangeTop_(), RangeButton_()).Borders.Color = color;
        }
        /// <summary>
        /// ボーダータイプ設定
        /// </summary>
        public void SetBorderType()
        {
            //sheet_.get_Range(RangeTop_(), RangeButton_()).Borders.LineStyle = Excel.XlLineStyle.xlDash;
            sheet_.get_Range(RangeTop_(), RangeButton_()).Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
        }
        /// <summary>
        /// ボーダータイプ設定
        /// </summary>
        /// <param name="border">BORDER_RANGEのタイプ指定</param>
        public void BorderRange(BORDER_RANGE border)
        {
            sheet_.Range[RangeTop_(), RangeButton_()].Borders[(Excel.XlBordersIndex)border].LineStyle = Excel.XlLineStyle.xlContinuous;
        }
        /// <summary>
        /// ボーダータイプ設定 + 色設定
        /// </summary>
        /// <param name="border">BORDER_RANGEのタイプ指定</param>
        /// <param name="color">カラー指定</param>
        public void BorderRange(BORDER_RANGE border, Color color)
        {
            sheet_.Range[RangeTop_(), RangeButton_()].Borders[(Excel.XlBordersIndex)border].LineStyle = Excel.XlLineStyle.xlContinuous;
            sheet_.Range[RangeTop_(), RangeButton_()].Borders[(Excel.XlBordersIndex)border].Color = color;
        }
        /// <summary>
        /// ボーダータイプ設定
        /// </summary>
        /// <param name="mat">EXCEL_MATRIXでの1行 or 1列指定</param>
        /// <param name="index">その行 or 列の場所指定</param>
        public void BorderRange(EXCEL_MATRIX mat, int index)
        {
            sheet_.Range[RangeTop_(mat, index), RangeButton_(mat, index)].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
        }
        /// <summary>
        /// ボーダータイプ指定
        /// </summary>
        /// <param name="mat">EXCEL_MATRIXでの1行 or 1列指定</param>
        /// <param name="index">その行 or 列の場所指定</param>
        /// <param name="color">カラー指定</param>
        public void BorderRange(EXCEL_MATRIX mat, int index, Color color)
        {
            sheet_.Range[RangeTop_(mat, index), RangeButton_(mat, index)].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            sheet_.Range[RangeTop_(mat, index), RangeButton_(mat, index)].Borders.Color = color;
        }
        #endregion
        /// <summary>
        /// シートを切り替える
        /// </summary>
        /// <param name="_change_index">指定するシートインデックス</param>
        public void ChangeSheet(int _change_index)
        {
            sheet_ = (Excel.Worksheet)book_.Worksheets[_change_index];
        }
        /// <summary>
        /// 全て縦書きに変更
        /// これをやるとセル範囲が自動調整されますのでSetColumnAutoFitは必要ありません
        /// </summary>
        public void SellxlVertical()
        {
            sheet_.get_Range(RangeTop_(), RangeButton_()).Orientation = Excel.XlOrientation.xlVertical;
        }
        /// <summary>
        /// 指定した行 or 列を縦書きに変更
        /// </summary>
        /// <param name="mat">EXCEL_MATRIXでの1行 or 1列指定</param>
        /// <param name="index">その行 or 列の場所指定</param>
        public void SellxlVertical(EXCEL_MATRIX mat, int index)
        {
            sheet_.get_Range(RangeTop_(mat, index), RangeButton_(mat, index)).Orientation = Excel.XlOrientation.xlVertical;
        }
        /// <summary>
        /// 全て横書きに変更 (デフォルトで横書きになっている場合は必要ない)
        /// </summary>
        public void SellxlHorizontal()
        {
            sheet_.get_Range(RangeTop_(), RangeButton_()).Orientation = Excel.XlOrientation.xlHorizontal;
        }
        /// <summary>
        /// 指定した行 or 列を横書きに変更
        /// </summary>
        /// <param name="mat">EXCEL_MATRIXでの1行 or 1列指定</param>
        /// <param name="index">その行 or 列の場所指定</param>
        public void SellxlHorizontal(EXCEL_MATRIX mat, int index)
        {
            sheet_.get_Range(RangeTop_(mat, index), RangeButton_(mat, index)).Orientation = Excel.XlOrientation.xlHorizontal;
        }
        /// <summary>
        /// 全ての文字が回転します
        /// </summary>
        /// <param name="rot">回転角度 -90 ～ 90で指定</param>
        public void SellxOrientation(int rot)
        {
            sheet_.Range[RangeTop_(), RangeButton_()].Orientation = rot;
        }
        /// <summary>
        /// 指定した行 or 列の文字を回転させます
        /// </summary>
        /// <param name="rot">回転角度 -90 ～ 90</param>
        /// <param name="mat">EXCEL_MATRIXでの1行 or 1列指定</param>
        /// <param name="index">その行 or 列の場所指定</param>
        /// <param name="top">その行 or 列の先頭指定</param>
        /// <param name="bottom">その行 or 列の末尾指定</param>
        public void SellxOrientation(int rot, EXCEL_MATRIX mat, int index, int top, int bottom)
        {
            sheet_.Range[RangeTop_(mat, index, top), RangeButton_(mat, index, bottom)].Orientation = rot;
        }
        /// <summary>
        /// 指定した行 or 列の文字を回転させます
        /// </summary>
        /// <param name="rot">回転角度 -90 ～ 90</param>
        /// <param name="mat">EXCEL_MATRIXでの1行 or 1列指定</param>
        /// <param name="index">その行 or 列の場所指定</param>
        public void SellxOrientation(int rot, EXCEL_MATRIX mat, int index)
        {
            sheet_.Range[RangeTop_(mat, index), RangeButton_(mat, index)].Orientation = rot;
        }
        /// <summary>
        /// 指定したセルの文字を回転させます
        /// </summary>
        /// <param name="rot">回転角度 -90 ～ 90</param>
        /// <param name="y">yのインデックス</param>
        /// <param name="x">xのインデックス</param>
        public void SellxOrientation(int rot, int y, int x)
        {
            sheet_.Range[RangeCell_(y, x - 1)].Orientation = rot;
        }
        /// <summary>
        /// 全ての高さを自動調整する
        /// </summary>
        public void SetWidthAutoFit()
        {
            sheet_.get_Range(RangeTop_(), RangeButton_()).Columns.AutoFit();
        }
        /// <summary>
        /// 全ての幅を自動調整する
        /// </summary>
        public void SetHeightAutoFit()
        {
            sheet_.get_Range(RangeTop_(), RangeButton_()).Rows.AutoFit();
        }
        /// <summary>
        /// セルサイズを調整する
        /// </summary>
        /// <param name="mat">EXCEL_MATRIXでの1行 or 1列指定</param>
        /// <param name="index">その行 or 列の場所指定</param>
        /// <param name="size">サイズ指定 0 ～ 255?</param>
        public void CellSizeWidth(EXCEL_MATRIX mat, int index, int size)
        {
            sheet_.get_Range(RangeTop_(mat, index), RangeButton_(mat, index)).Cells.ColumnWidth = size;
        }
        /// <summary>
        /// セルサイズを調整する
        /// </summary>
        /// <param name="size">サイズ指定 0 ～ 255?</param>
        public void CellSizeWidth(int size)
        {
            sheet_.get_Range(RangeTop_(), RangeButton_()).Cells.ColumnWidth = size;
        }
        /// <summary>
        /// セルサイズを調整する
        /// </summary>
        /// <param name="mat">EXCEL_MATRIXでの1行 or 1列指定</param>
        /// <param name="index">その行 or 列の場所指定</param>
        /// <param name="size"></param>
        public void CellSizeHeight(EXCEL_MATRIX mat, int index, int size)
        {
            sheet_.get_Range(RangeTop_(mat, index), RangeButton_(mat, index)).Cells.Rows.RowHeight = size;
        }
        /// <summary>
        /// 表示するエクセルのズーム倍率設定
        /// </summary>
        /// <param name="zoom">ズーム指定 0 ～400 (デフォルト100)</param>
        public void SetZoom(int zoom)
        {
            excel_.Application.ActiveWindow.Zoom = zoom;
        }
        #region 太字
        /// <summary>
        /// 太字設定
        /// </summary>
        /// <param name="mat">EXCEL_MATRIXでの1行 or 1列指定</param>
        /// <param name="index">その行 or 列の場所指定行</param>
        public void SetBold(EXCEL_MATRIX mat, int index)
        {
            sheet_.Range[RangeTop_(mat, index), RangeButton_(mat, index)].Font.Bold = true;
        }
        /// <summary>
        /// 太字設定
        /// </summary>
        /// <param name="mat">EXCEL_MATRIXでの1行 or 1列指定</param>
        /// <param name="index">その行 or 列の場所指定</param>
        /// <param name="color">カラー指定</param>
        public void SetBold(EXCEL_MATRIX mat, int index, Color color)
        {
            sheet_.Range[RangeTop_(mat, index), RangeButton_(mat, index)].Font.Bold = true;
            sheet_.Range[RangeTop_(mat, index), RangeButton_(mat, index)].Font.Color = color;
        }
        /// <summary>
        /// 太字設定 (セル指定)
        /// </summary>
        /// <param name="y">yのインデックス</param>
        /// <param name="x">xのインデックス</param>
        public void SetBold(int y, int x)
        {
            sheet_.Range[RangeCell_(y, x)].Font.Bold = true;
        }
        /// <summary>
        /// 太字設定 
        /// </summary>
        /// <param name="y">yのインデックス</param>
        /// <param name="x">xのインデックス</param>
        /// <param name="size">文字サイズ指定</param>
        public void SetBold(int y, int x, int size)
        {
            sheet_.Range[RangeCell_(y, x)].Font.Bold = true;
            sheet_.Range[RangeCell_(y, x)].Font.Size = size;
        }
        /// <summary>
        /// 太字設定
        /// </summary>
        /// <param name="y">yのインデックスy</param>
        /// <param name="x">xのインデックス</param>
        /// <param name="color">カラー指定</param>
        public void SetBold(int y, int x, Color color)
        {
            sheet_.Range[RangeCell_(y, x)].Font.Bold = true;
            sheet_.Range[RangeCell_(y, x)].Font.Color = color;
        }
        #endregion
        #region フォントカラー
        /// <summary>
        /// フォントカラー
        /// </summary>
        /// <param name="color">カラー指定</param>
        public void SetFontColor(Color color)
        {
            sheet_.Range[RangeTop_(), RangeButton_()].Font.Color = System.Drawing.ColorTranslator.ToOle(color);
        }
        /// <summary>
        /// フォントカラー
        /// </summary>
        /// <param name="mat">EXCEL_MATRIXでの1行 or 1列指定</param>
        /// <param name="index">その行 or 列の場所指定その行 or 列の場所指定</param>
        /// <param name="color">カラー指定</param>
        public void SetFontColor(EXCEL_MATRIX mat, int index, Color color)
        {
            sheet_.Range[RangeTop_(mat, index), RangeButton_(mat, index)].Font.Color = System.Drawing.ColorTranslator.ToOle(color);
        }
        /// <summary>
        /// フォントカラー
        /// </summary>
        /// <param name="y">yのインデックス</param>
        /// <param name="x">xのインデックス</param>
        /// <param name="color">カラー指定</param>
        public void SetFontColor(int y, int x, Color color)
        {
            sheet_.Range[RangeCell_(y, x)].Font.Color = System.Drawing.ColorTranslator.ToOle(color);
        }
        #endregion
        #region セル背景カラー
        /// <summary>
        /// セルカラー設定
        /// </summary>
        /// <param name="color">カラー指定</param>
        public void SetInteriorColor(Color color)
        {
            sheet_.Range[RangeTop_(), RangeButton_()].Interior.Color = System.Drawing.ColorTranslator.ToOle(color);
        }
        /// <summary>
        /// セルカラー設定
        /// </summary>
        /// <param name="mat">EXCEL_MATRIXでの1行 or 1列指定</param>
        /// <param name="index">その行 or 列の場所指定</param>
        /// <param name="color">カラー指定</param>
        public void SetInteriorColor(EXCEL_MATRIX mat, int index, Color color)
        {
            sheet_.Range[RangeTop_(mat, index), RangeButton_(mat, index)].Interior.Color = System.Drawing.ColorTranslator.ToOle(color);
        }
        /// <summary>
        /// セルカラー設定
        /// </summary>
        /// <param name="y"></param>
        /// <param name="x"></param>
        /// <param name="color">カラー指定</param>
        public void SetInteriorColor(int y, int x, Color color)
        {
            sheet_.Range[RangeCell_(y, x)].Interior.Color = System.Drawing.ColorTranslator.ToOle(color);
        }
        #endregion
        #region フォント名
        /// <summary>
        /// フォント名設定
        /// </summary>
        /// <param name="font_name">フォント名</param>
        public void SetFontName(string font_name)
        {
            sheet_.Range[RangeTop_(), RangeButton_()].Font.Name = font_name;
        }
        /// <summary>
        /// フォント名設定
        /// </summary>
        /// <param name="mat">EXCEL_MATRIXでの1行 or 1列指定</param>
        /// <param name="index">その行 or 列の場所指定</param>
        /// <param name="font_name">フォント名</param>
        public void SetFontName(EXCEL_MATRIX mat, int index, string font_name)
        {
            sheet_.Range[RangeTop_(mat, index), RangeButton_(mat, index)].Font.Name = font_name;
        }
        /// <summary>
        /// フォント名設定
        /// </summary>
        /// <param name="y"></param>
        /// <param name="x"></param>
        /// <param name="font_name">フォント名</param>
        public void SetFontName(int y, int x, string font_name)
        {
            sheet_.Range[RangeCell_(y, x)].Font.Name = font_name;
        }
        #endregion

        #endregion
        #region 印刷
        /// <summary>
        /// PDF作成
        /// </summary>
        /// <param name="file_xls">生成するファイル名(フルパス)</param>
        /// <param name="top_y"></param>
        /// <param name="botton_y"></param>
        public void CreatePDF(string file_xls, int top_y, int botton_y)
        {
            string file_pdf = Path.ChangeExtension(file_xls, "pdf");
            var s = new string[] { RangeTop_(EXCEL_MATRIX.X, 0, top_y), RangeButton_(EXCEL_MATRIX.X, cell_max_y_ - 1, botton_y) };

            sheet_.Range[s[0], s[1]].ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF,
                file_pdf,
                Excel.XlFixedFormatQuality.xlQualityStandard,
                true,
                true,
                Type.Missing,
                Type.Missing,
                false,
                Type.Missing);
        }
        /// <summary>
        /// PDF作成
        /// </summary>
        /// <param name="file_xls">生成するファイル名(フルパス)</param>
        public void CreatePDF(string file_xls)
        {
            string file_pdf = Path.ChangeExtension(file_xls, "pdf");
            book_.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF,
                file_pdf,
                Excel.XlFixedFormatQuality.xlQualityStandard,
                true,
                true,
                Type.Missing,
                Type.Missing,
                false,
                Type.Missing);
        }
        /// <summary>
        /// PDF作成
        /// </summary>
        /// <param name="file_xls">生成するファイル名(フルパス)</param>
        /// <param name="open">表示するかどうか</param>
        public void CreatePDF(string file_xls, bool open)
        {
            string file_pdf = Path.ChangeExtension(file_xls, "pdf");
            book_.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF,
                file_pdf,
                Excel.XlFixedFormatQuality.xlQualityStandard,
                true,
                true,
                Type.Missing,
                Type.Missing,
                open,
                Type.Missing);
        }
        /// <summary>
        /// 印刷のグリッドライン設定？
        /// </summary>
        public void PrintGridLine()
        {
            sheet_.PageSetup.PrintGridlines = true;
        }
        /// <summary>
        /// 印刷の向き設定 現在横
        /// </summary>
        public void PrintOrientation()
        {
            sheet_.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
        }
        /// <summary>
        /// 印刷のサイズ設定
        /// </summary>
        /// <param name="size">PRINT_SIZEによるサイズ指定</param>
        public void PrintSize(PRINT_SIZE size)
        {
            sheet_.PageSetup.PaperSize = (Excel.XlPaperSize)size;
        }
        /// <summary>
        /// 印刷の範囲を設定します 
        /// </summary>
        public void PrintArea()
        {
            sheet_.PageSetup.PrintArea = RangeTop_() + ":" + RangeButton_();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="range_x"></param>
        public void PrintAreaX(int range_x)
        {
            sheet_.PageSetup.PrintArea = RangeTop_(EXCEL_MATRIX.X, range_x, 0) + ":" + RangeButton_();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="top_y"></param>
        /// <param name="botton_y"></param>
        public void PrintAreaY(int top_y, int botton_y)
        {
            sheet_.PageSetup.PrintArea = RangeTop_(EXCEL_MATRIX.X, 0, top_y) + ":" + RangeButton_(EXCEL_MATRIX.X, cell_max_y_ - 1, botton_y);
        }
        /// <summary>
        /// エクセルデータの枠を追加します ("A1"などの枠)
        /// </summary>
        public void Printhoge()
        {
            sheet_.PageSetup.PrintHeadings = true;
        }
        /// <summary>
        /// 中間設定
        /// </summary>
        /// <param name="margin"></param>
        public void PrintMarginCenter(int margin)
        {
            sheet_.PageSetup.LeftMargin = margin;
            sheet_.PageSetup.RightMargin = margin;
            sheet_.PageSetup.TopMargin = margin;
            sheet_.PageSetup.BottomMargin = margin;
        }
        /// <summary>
        /// 中央寄せ設定
        /// </summary>
        public void PrintMiddle()
        {
            //水平方向中央寄せ
            sheet_.PageSetup.CenterHorizontally = true;
            //垂直方向中央寄せ
            sheet_.PageSetup.CenterVertically = true;
        }
        #region HederFooter デリゲート
        delegate void HeaderFooterDelegate(PRINT_POSITION pos, string s);
        delegate void HeaderFooterIntDelegate(int pos, string s);
        delegate void HeaderFooterDelegate_(string s);
        /// <summary>
        /// ヘッダー・フッター設定
        /// </summary>
        /// <param name="hf"></param>
        /// <param name="pos"></param>
        /// <param name="s"></param>
        public void PrintHeaderFooter(PRINT_HEAD_FOOT hf, PRINT_POSITION pos, string s)
        {
            HeaderFooterDelegate[] fd = 
            {
                PrintHeader,
                PrintFooter,
            };
            fd[(int)hf](pos, s);
        }
        /// <summary>
        /// ヘッダー・フッター設定
        /// </summary>
        /// <param name="pos"></param>
        /// <param name="s"></param>
        public void PrintFooter(PRINT_POSITION pos, string s)
        {
            HeaderFooterDelegate_[] fd = 
            {
                PrintFooterLeft_,
                PrintFooterCenter_,
                PrintFooterRight_,
            };
            fd[(int)pos](s);
        }
        /// <summary>
        /// ヘッダー・フッター設定
        /// </summary>
        /// <param name="pos"></param>
        /// <param name="s"></param>
        public void PrintHeader(PRINT_POSITION pos, string s)
        {
            HeaderFooterDelegate_[] hd = 
            {
                PrintHeaderLeft_,
                PrintHeaderCenter_,
                PrintHeaderRight_,
            };
            hd[(int)pos](s);
        }
        /// <summary>
        /// ヘッダー・フッター設定
        /// 0, 1, 2
        /// 3, 4, 5 配置
        /// </summary>
        /// <param name="index"></param>
        /// <param name="s"></param>
        public void PrintHeaderFooter(int index, string s)
        {
            HeaderFooterIntDelegate[] fd = 
            {
                PrintHeader,
                PrintFooter,
            };
            fd[index / 3](index % 3, s);
        }
        /// <summary>
        /// ヘッダー・フッター設定
        /// </summary>
        /// <param name="hf"></param>
        /// <param name="pos"></param>
        /// <param name="s"></param>
        public void PrintHeaderFooter(int hf, int pos, string s)
        {
            HeaderFooterIntDelegate[] fd = 
            {
                PrintHeader,
                PrintFooter,
            };
            fd[hf](pos, s);
        }
        /// <summary>
        /// フッター設定
        /// </summary>
        /// <param name="pos"></param>
        /// <param name="s"></param>
        public void PrintFooter(int pos, string s)
        {
            HeaderFooterDelegate_[] fd = 
            {
                PrintFooterLeft_,
                PrintFooterCenter_,
                PrintFooterRight_,
            };
            fd[pos](s);
        }
        /// <summary>
        /// ヘッダー設定
        /// </summary>
        /// <param name="pos"></param>
        /// <param name="s"></param>
        public void PrintHeader(int pos, string s)
        {
            HeaderFooterDelegate_[] hd = 
            {
                PrintHeaderLeft_,
                PrintHeaderCenter_,
                PrintHeaderRight_,
            };
            hd[pos](s);
        }
        /// <summary>
        /// フッター左設定
        /// </summary>
        /// <param name="s"></param>
        private void PrintFooterLeft_(string s)
        {
            sheet_.PageSetup.LeftFooter = s;
        }
        /// <summary>
        /// フッター右設定
        /// </summary>
        /// <param name="s"></param>
        private void PrintFooterRight_(string s)
        {
            sheet_.PageSetup.RightFooter = s;
        }
        /// <summary>
        /// フッター真ん中設定
        /// </summary>
        /// <param name="s"></param>
        private void PrintFooterCenter_(string s)
        {
            sheet_.PageSetup.CenterFooter = s;
        }
        /// <summary>
        /// ヘッダー左設定
        /// </summary>
        /// <param name="s"></param>
        private void PrintHeaderLeft_(string s)
        {
            sheet_.PageSetup.LeftHeader = s;
        }
        /// <summary>
        ///  ヘッダー右設定
        /// </summary>
        /// <param name="s"></param>
        private void PrintHeaderRight_(string s)
        {
            sheet_.PageSetup.RightHeader = s;
        }
        /// <summary>
        /// ヘッダー真ん中設定
        /// </summary>
        /// <param name="s"></param>
        private void PrintHeaderCenter_(string s)
        {
            sheet_.PageSetup.CenterHeader = s;
        }
        #endregion
        #endregion
        #region ピクチャ設定
        /// <summary>
        /// ピクチャをエクセルに表示
        /// </summary>
        /// <param name="picture_path"></param>
        public void Picture(string picture_path)
        {
            sheet_.Shapes.AddPicture(picture_path,
                    Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 10, 10, 10, 10);
        }
        /// <summary>
        /// ピクチャをエクセルに表示
        /// 画像をセル範囲に入るように拡大 (微調整)
        /// </summary>
        /// <param name="picture_path"></param>
        /// <param name="start_y"></param>
        /// <param name="start_x"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        public void PictureExpantion(string picture_path, int start_y, int start_x, float width, float height)
        {
            float _width = (float)sheet_.Range[RangeCell_(0, 0)].Width;
            float _height = (float)sheet_.Range[RangeCell_(0, 0)].RowHeight;

            int w = (int)(width / _width) + 1;
            int h = (int)(height / _height) + 1;
            sheet_.Shapes.AddPicture(picture_path,
                    Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, sheet_.Range[RangeCell_(start_y, start_x)].Left, sheet_.Range[RangeCell_(start_y, start_x)].Top, sheet_.Range[RangeCell_(h, w)].Left, sheet_.Range[RangeCell_(h, w)].Top);
        }
        /// <summary>
        /// ピクチャをエクセルに表示
        /// 画像をセル範囲に入るように収縮 (微調整)
        /// </summary>
        /// <param name="picture_path"></param>
        /// <param name="start_y"></param>
        /// <param name="start_x"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        public void PictureContraction(string picture_path, int start_y, int start_x, float width, float height)
        {
            float _width = (float)sheet_.Range[RangeCell_(0, 0)].Width;
            float _height = (float)sheet_.Range[RangeCell_(0, 0)].RowHeight;

            int w = (int)(width / _width);
            int h = (int)(height / _height);
            sheet_.Shapes.AddPicture(picture_path,
                    Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, sheet_.Range[RangeCell_(start_y, start_x)].Left, sheet_.Range[RangeCell_(start_y, start_x)].Top, sheet_.Range[RangeCell_(h, w)].Left, sheet_.Range[RangeCell_(h, w)].Top);
        }
        /// <summary>
        /// ピクチャをエクセルに表示
        /// </summary>
        /// <param name="picture_path"></param>
        /// <param name="left"></param>
        /// <param name="top"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        public void Picture(string picture_path, float left, float top, float width, float height)
        {
            sheet_.Shapes.AddPicture(picture_path,
                    Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, left, top, width, height);
        }
        #endregion
        #region 条件式指定
        /// <summary>
        /// 条件付き書式設定
        /// ※汎用性がないため、使用しないでください 
        /// ※セーブする際、.xlsx形式でないと反映されません。
        /// </summary>
        /// <param name="mat">EXCEL_MATRIXでの1行 or 1列指定</param>
        /// <param name="index1"></param>
        /// <param name="index2"></param>
        /// <param name="color">カラー指定</param>
        public void ConditionNotEqual(EXCEL_MATRIX mat, int index1, int index2, Color color)
        {
            // 数値をアルファベットに変換
            string[] cell = { GetColumnName_(index1), GetColumnName_(index2) };
            // 条件付き書式の範囲
            string[] range = { cell[0] + (2).ToString(), cell[0] + (cell_max_y_).ToString() };
            // 判定(比較)
            string range_condition = "=$" + cell[0] + (2).ToString() + "<>$" + cell[1] + (2).ToString();

            Excel.FormatCondition condition =
                (Excel.FormatCondition)sheet_.Range[range[0], range[1]].FormatConditions.Add
                    (
                        Excel.XlFormatConditionType.xlExpression, Type.Missing, range_condition
                    );

            condition.Interior.Color = color;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(condition);
        }
        /// <summary>
        /// 同じ文字の色替え
        /// </summary>
        /// <param name="mat">EXCEL_MATRIXでの1行 or 1列指定</param>
        /// <param name="index">その行 or 列の場所指定</param>
        /// <param name="color">カラー指定</param>
        public void ConditionOverlap(EXCEL_MATRIX mat, int index, Color color)
        {
            // 数値をアルファベットに変換
            string cell = GetColumnName_(index);
            // 条件付き書式の範囲
            string[] range = { cell + (2).ToString(), cell + (cell_max_y_).ToString() };
            // 判定 "=COUNTIF($G$2:$G$100,G2)>1");
            var range_condition = "=COUNTIF($" + cell + "$" + (2).ToString() + ":$" + cell + "$" + (cell_max_y_).ToString() + "," + cell + (2).ToString() + ")>1";

            Excel.FormatCondition condition =
                (Excel.FormatCondition)sheet_.Range[range[0], range[1]].FormatConditions.Add
                    (
                        Excel.XlFormatConditionType.xlExpression, Type.Missing, range_condition
                    );
            condition.Interior.Color = color;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(condition);
        }
        #endregion
        #endregion
    }
}