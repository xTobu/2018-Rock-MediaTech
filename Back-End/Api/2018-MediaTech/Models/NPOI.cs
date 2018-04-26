using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace _2018_MediaTech.Models
{
    public class NPOI
    {
        public HSSFWorkbook NPOIGridviewToExcel(DataTable dt)
        {//製作成EXCEL

            int sheetCount = 1; int sheetRowIdx = 65535; int currentRowIdx = 1;
            HSSFWorkbook workBook = new HSSFWorkbook();
            ISheet sheet = SheetCreate(workBook, sheetCount, dt);
            for (int i = 0; i < dt.Rows.Count; i++)//塞值
            {
                if (currentRowIdx >= sheetRowIdx)
                {
                    sheetCount++; currentRowIdx = 1;
                    sheet = SheetCreate(workBook, sheetCount, dt);
                }
                IRow rowtemp = sheet.CreateRow(currentRowIdx++);
                for (int j = 0; j < dt.Columns.Count; j++)
                    rowtemp.CreateCell(j).SetCellValue(dt.Rows[i][j].ToString());
            }
            //MemoryStream ms = new MemoryStream();
            //workBook.Write(ms);
            ////Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + HttpUtility.UrlEncode(excel_name, System.Text.Encoding.UTF8) + ".xls").ToString());
            ////Response.BinaryWrite(ms.ToArray());
            //workBook = null;
            //ms.Close();
            //ms.Dispose();
            return workBook;
        }

        public ISheet SheetCreate(HSSFWorkbook workBook, int sheetCount, DataTable dt)
        {
            ISheet sheet = workBook.CreateSheet("sheet" + sheetCount.ToString());
            IRow row = sheet.CreateRow(0);//Title
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                row.CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
            }
            return sheet;
        }
    }
}