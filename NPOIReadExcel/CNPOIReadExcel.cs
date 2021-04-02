using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NPOIReadExcel
{
    public class CNPOIReadExcel
    {
        private IWorkbook mWk = null;
        private ISheet    mSheet = null;  
        public bool OpenExcel(string fileName, ref string errInfo) {

            if (!System.IO.File.Exists(fileName))
            {
                errInfo = "文件不存在！";
                return false;
            }
            string extension = System.IO.Path.GetExtension(fileName);
             
            try
            {
                FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                //FileStream fs = File.OpenRead(fileName);
                if (extension.Equals(".xls"))
                {
                    //把xls文件中的数据写入wk中
                    mWk = new HSSFWorkbook(fs); 
                }
                else
                {
                    //把xlsx文件中的数据写入wk中
                    mWk = new XSSFWorkbook(fs); 
                }

                mSheet = mWk.GetSheetAt(0);

                fs.Close();
            }
            catch (Exception ex)
            {
                errInfo = ex.Message;
                return false;
            }

            return true;
        }


        public string GetCell( int row,int col ) {
            string cellValue = "";
            IRow iRow = mSheet.GetRow(row);
            if (iRow == null)
                return "";

            ICell iCell = iRow.GetCell(col);
            if (iCell == null)
                return "";

            CellType cType = iCell.CellType; // 获取单元格中的类型
            //判断当前单元格的数据类型 
            switch (cType)
            {
                case CellType.Numeric: //数字 
                    {
                        cellValue = iCell.NumericCellValue.ToString();
                    }
                    break;
                case CellType.String: //字符串 
                    {
                        cellValue = iCell.StringCellValue;
                    }
                    break;
                case CellType.Boolean: // 布尔
                    {
                        cellValue = iCell.BooleanCellValue.ToString();
                    }
                    break;
                case CellType.Formula: // 公式
                    {
                        cellValue = iCell.NumericCellValue.ToString();
                    }
                    break;
            } 
            return cellValue; 
        }
         
        public bool CloseExcel()
        {
            mWk.Close();
            return true;
        }
    }
}
