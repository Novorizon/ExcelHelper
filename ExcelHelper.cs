using System.Collections.Generic;
using NPOI.SS.UserModel;
using System.IO;
using System;
using UnityEngine;
using NPOI.XSSF.UserModel;

namespace LevelEditor
{
    public enum ExcelValue
    {
        Int = 0,
        Double = 1,
        String = 2,
        Bool = 3,
        Date = 4
    }

    public class ExcelHelper
    {
        ExcelReader excelReader;
        public bool success = false;
        private string path;
        private string sheet;

        private XSSFWorkbook workbook;
        private FileStream address;
        private ISheet Sheet;
        //private XSSFCell Cell;
        private Dictionary<int, List<string>> cells = new Dictionary<int, List<string>>();
        private List<List<string>> Values = new List<List<string>>();

        public ExcelHelper()
        {
        }
        public ExcelHelper(string path, FileMode fileMode = FileMode.Open, FileAccess FileAccess = FileAccess.Read, FileShare fileShare = FileShare.Read, string sheet = "")
        {
            this.path = path;
            this.sheet = sheet;

            try
            {
                address = new FileStream(path, fileMode, FileAccess, FileShare.ReadWrite);
                address.Position = 0;
                if (fileMode == FileMode.Open)
                {
                    workbook = new XSSFWorkbook(address);
                    if (sheet == "")
                        Sheet = workbook.GetSheetAt(0);
                    else
                        Sheet = workbook.GetSheet(sheet);
                }
                else
                {
                    workbook = new XSSFWorkbook();
                    Sheet = (XSSFSheet)workbook.CreateSheet("Sheet1");
                }


                success = true;
                address.Close();
            }
            catch (Exception e)
            {
                success = false;
                Console.WriteLine(e.Message);
                return;
            }
        }

        public void Open(string path, string sheetName = "Sheet1")
        {
            this.path = path;
            this.sheet = sheetName;
            excelReader = new ExcelReader(path, sheetName, ExcelReader.Mode.Open);
        }
        public void Create(string path, string sheetName = "Sheet1")
        {
            this.path = path;
            this.sheet = sheetName;
            excelReader = new ExcelReader(path, sheetName, ExcelReader.Mode.Create);
        }

        public int RowMax { get { return Sheet.LastRowNum; } }
        public string SheetName { get { return Sheet.SheetName; } }
        public XSSFRow CreateRow(int row, int cellCount = 0) => excelReader.CreateRow(row, cellCount);
        public XSSFRow GetRow(int row) => excelReader.GetRow(row);
        public XSSFCell GetCell(int row, int column) => excelReader.GetCell(row, column);
        public XSSFCell CreateCell(int row, int column) => excelReader.CreateCell(row, column);

        public void RemoveRow(int row, int count = 1) => excelReader.RemoveRow(row, count);
        public void DeleteRow(int startRow, int endRow) => excelReader.DeleteRow(startRow, endRow);


        public void CreateValue(int row, int column, int value) => excelReader.CreateValue(row, column, value);
        public void CreateValue(int row, int column, string value) => excelReader.CreateValue(row, column, value);
        public void CreateValue(int row, int column, double value) => excelReader.CreateValue(row, column, value);
        public void CreateValue(int row, int column, bool value) => excelReader.CreateValue(row, column, value);
        public void CreateValue(int row, int column, DateTime value) => excelReader.CreateValue(row, column, value);

        public void UpdateValue(int row, int column, string value) => excelReader.UpdateValue(row, column, value);
        public void UpdateValue(int row, int column, int value) => excelReader.UpdateValue(row, column, value);
        public void UpdateValue(int row, int column, double value) => excelReader.UpdateValue(row, column, value);
        public void UpdateValue(int row, int column, DateTime value) => excelReader.UpdateValue(row, column, value);
        public void UpdateValue(int row, int column, bool value) => excelReader.UpdateValue(row, column, value);

        public string GetValue(int row, int column) => excelReader.GetValue(row, column);
        public int GetValueInt(int row, int column) => excelReader.GetValueInt(row, column);
        public Double GetValueDouble(int row, int column) => excelReader.GetValueDouble(row, column);
        public String GetValueString(int row, int column) => excelReader.GetValueString(row, column);
        public bool GetValueBool(int row, int column) => excelReader.GetValueBool(row, column);
        public DateTime GetValueDate(int row, int column) => excelReader.GetValueDate(row, column);

        public List<string> GetRowValue(int row, int columnStart)
        {
            XSSFRow Row = GetRow(row);
            List<string> values = new List<string>();
            for (int i = columnStart; i <= Row.LastCellNum; i++)
            {
                string value = GetValue(row, i);
                values.Add(value);
            }

            return values;
        }


        public List<List<string>> Read() => excelReader.Read();
        public List<List<int>> ReadInt() => excelReader.ReadInt();
        public List<List<double>> ReadDouble() => excelReader.ReadDouble();


        public bool Write(List<List<string>> Values) => excelReader.Write(Values);
        public bool Write(int row, int column, string vlaue) => excelReader.Write(row, column, vlaue);


        public void Write()
        {

            try
            {
                excelReader.Write();


            }
            catch (Exception e)
            {
                Debug.LogError(e.Message);
            }
        }
        public Dictionary<int, List<string>> GetAllValue(int row = 0, int column = 0)
        {
            if (Sheet == null)
                return null;

            XSSFRow Row = GetRow(row);
            cells.Clear();

            for (int i = row; i <= Sheet.LastRowNum; i++)
            {
                List<string> values = new List<string>();
                for (int j = column; j < Row.LastCellNum; j++)
                {
                    var value = GetCell(i, j)?.ToString();
                    values.Add(value);
                }
                cells.Add(i, values);
            }


            return cells;
        }




        //根据自己需求改
        //for (int i = 0; i < hangNum.Length; i++)
        //{
        //    HSSFRow Row = (HSSFRow)sheet01.CreateRow((short)i);//为工作表定义行 
        //    HSSFCell cell = (HSSFCell)Row.CreateCell((short)0);//为第i行  定义列
        //    cell.SetCellValue(hangNum[i]);//给第i列添加数值
        //    if (i < hang.Length)
        //    {
        //        HSSFCell cell02 = (HSSFCell)Row.CreateCell((short)1);
        //        cell02.SetCellValue(hang[i]);
        //    }
        //    else
        //    {
        //        HSSFCell cell02 = (HSSFCell)Row.CreateCell((short)1);
        //        cell02.SetCellValue("");
        //    }
        //    #region[格式设置]
        //    //Row.RowStyle = MyWorkbook.CreateCellStyle();//定义行样式
        //    //Row.RowStyle.BorderBottom = BorderStyle.Double;//更改行边界
        //    //cell.CellStyle = MyWorkbook.CreateCellStyle();//定义单元格格式
        //    //cell.CellStyle.BorderRight = BorderStyle.Thin;//改变一小格边界
        //    //cell.CellStyle.BorderBottom = BorderStyle.Dashed;
        //    //cell.CellStyle.BottomBorderColor = HSSFColor.Red.Index;

        //    //HSSFFont MyFont = (HSSFFont)MyWorkbook.CreateFont();//定义字体
        //    改变字体、字体高度、字体颜色、eto
        //    //MyFont.FontName = "Tahoma";
        //    //MyFont.FontHeightInPoints = 14;
        //    //MyFont.Color = HSSFColor.Gold.Index;
        //    //MyFont.Boldweight = (short)FontBoldWeight.Bold;

        //    //设置单元格字体
        //    //cell.CellStyle.SetFont(MyFont);
        //    #endregion
        //}
        //}

        //关闭打开的excel
        public void Close()
        {
            Sheet = null;
            //workbook?.Close();
            address?.Close();
        }
    }
}
