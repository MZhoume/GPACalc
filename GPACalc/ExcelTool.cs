using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Cells;
using System.Data;

namespace GPACalc
{
    public class ExcelTool
    {
        private static string outFileName = "";
        private static string inFileName = "";
        
        /// <summary>
        /// 保存Workbook表到外部Excel文件
        /// </summary>
        /// <param name="book">需要保存的Workbook</param>
        /// <param name="outfilename">保存的路径</param>
        /// <returns>是否保存成功</returns>
        public static bool GenerateOuterExcel(Workbook book,string outfilename)
        {
            outFileName = outfilename;
            try
            {
                book.Save(outFileName);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 从外部打开Excel文件并返回Workbook
        /// </summary>
        /// <param name="infilename">打开文件的路径</param>
        /// <returns>返回的Workbook</returns>
        public static Workbook LoadOuterExcel(string infilename)
        {
            inFileName = infilename;
            try
            {
                Workbook book = new Workbook(inFileName);
                return book;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 在Workbook中添加标题
        /// </summary>
        /// <param name="book">需添加标题的Workbook</param>
        /// <param name="title">添加的标题</param>
        /// <param name="NeedTime">是否需要添加时间</param>
        /// <param name="columnCount">标题占的列数</param>
        /// <returns>成功会返回Workbook，失败返回null</returns>
        private static Workbook AddTitle(Workbook book, string title,bool NeedTime, int columnCount)
        {
            try
            {
                Worksheet sheet = book.Worksheets[0];
                sheet.Cells.Merge(0, 0, 1, columnCount);
                Cell cell1 = sheet.Cells[0, 0];
                cell1.PutValue(title);
                Style cell1style = cell1.GetStyle();
                cell1style.HorizontalAlignment = TextAlignmentType.Center;
                cell1style.Font.Name = "黑体";
                cell1style.Font.Size = 14;
                cell1style.Font.IsBold = true;
                cell1.SetStyle(cell1style);
                if (NeedTime)
                {
                    sheet.Cells.Merge(1, 0, 1, columnCount);
                    Cell cell2 = sheet.Cells[1, 0];
                    cell2.PutValue("创建时间：" + DateTime.Now.ToLocalTime());
                    cell2.SetStyle(cell1style);
                }
                return book;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 从Datatable中读取列名添加到Workbook中某行（从0开始计数）
        /// </summary>
        /// <param name="book">添加到的Workbook</param>
        /// <param name="dt">需要读取的Datatable表</param>
        /// <param name="row">在第几行添加</param>
        /// <returns>成功会返回Workbook，失败返回null</returns>
        private static Workbook AddHeader(Workbook book, DataTable dt,int row)
        {
            try
            {
                Worksheet sheet = book.Worksheets[0];
                Cell cell = null;
                for (int col = 0; col < dt.Columns.Count; col++)
                {
                    cell = sheet.Cells[row, col];
                    cell.PutValue(dt.Columns[col].ColumnName);
                    Style cellstyle = cell.GetStyle();
                    cellstyle.Font.IsBold = true;
                    cellstyle.Font.Name = "黑体";
                    cell.SetStyle(cellstyle);
                }
                return book;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 从Datatable中读取表体添加到Workbook中(从0开始计数)
        /// </summary>
        /// <param name="book">需要添加到的Workbook</param>
        /// <param name="dt">需要读取表体的Datatable</param>
        /// <param name="bookrow">从Workbook第几行开始添加</param>
        /// <param name="dtrow">从Datatable第几行开始读取</param>
        /// <returns>成功会返回Workbook，失败返回null</returns>
        private static Workbook AddBody(Workbook book,DataTable dt,int bookrow,int dtrow)
        {
            try
            {
                Worksheet sheet = book.Worksheets[0];
                for (int r = 0; r < (dt.Rows.Count - dtrow + 1); r++)
                {
                    for (int c = 0; c < dt.Columns.Count; c++)
                    {
                        sheet.Cells[r + bookrow, c].PutValue(dt.Rows[r + dtrow][c].ToString());
                    }
                }
                return book;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 从Datatable导出到外部Excel文件
        /// </summary>
        /// <param name="dt">需要导出的Datatable</param>
        /// <param name="outfilename">外部文件存储路径</param>
        /// <param name="sheetName">表名</param>
        /// <param name="title">表的标题</param>
        /// <param name="needtime">是否需要表的创建时间</param>
        /// <returns>是否成功</returns>
        public static bool DatatableToExcel(DataTable dt,string outfilename,string sheetName,string title,bool needtime)
        {
            bool createSuccess = false;
            Workbook book=new Workbook();
            Worksheet sheet = book.Worksheets[0];
            int headrow = 1;
            outFileName = outfilename;
            try
            {
                sheet.Name = sheetName;
                AddTitle(book,title,needtime,dt.Columns.Count);
                if (needtime)
                {
                    headrow = 2;
                }
                AddHeader(book,dt,headrow);
                AddBody(book,dt,headrow+1,0);
                sheet.AutoFitColumns();
                sheet.AutoFitRows();
                createSuccess = GenerateOuterExcel(book,outFileName);
                return createSuccess;
            }
            catch
            {
                return false; ;
            }
        }

        /// <summary>
        /// 从外部读取Excel文件返回Datatable
        /// </summary>
        /// <param name="infilename">外部Excel路径</param>
        /// <returns>读取返回的Datatable</returns>
        public static DataTable ExcelToDatatable(string infilename)
        {
            inFileName=infilename;
            Workbook book = LoadOuterExcel(inFileName);
            Worksheet sheet = book.Worksheets[0];
            Cells cells = sheet.Cells;
            DataTable dt_import=cells.ExportDataTableAsString(0,0,cells.MaxDataRow+1,cells.MaxDataColumn+1,false);
            for (int i = 0; i < dt_import.Columns.Count; i++)
            {
                dt_import.Columns[i].ColumnName = dt_import.Rows[0][i].ToString();
            }
            dt_import.Rows.Remove(dt_import.Rows[0]);
            return dt_import;
        }
    }
}
