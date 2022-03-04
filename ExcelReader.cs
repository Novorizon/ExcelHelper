using System.Collections.Generic;
using NPOI.SS.UserModel;
using System.IO;
using System;
using UnityEngine;
using NPOI.XSSF.UserModel;

namespace LevelEditor
{
    public class ExcelReader
    {
        public bool success = false;
        private string path;
        private string sheet;

        private XSSFWorkbook workbook;
        private ISheet Sheet;


        public enum Mode
        {
            Create = 1,
            Open = 2,
            Append = 3
        }

        public ExcelReader(string path, string sheetName = "Sheet1", Mode mode = Mode.Open)
        {
            this.path = path;
            this.sheet = sheetName;
            if (mode == Mode.Open)
            {
                Open();
            }
            else
            {
                Create();
            }
        }

        public void Create()
        {
            using (FileStream fs = File.OpenWrite(path))
            {
                workbook = new XSSFWorkbook();
                Sheet = workbook.CreateSheet(sheet);
                workbook.Write(fs);
            }
        }
        public void Open()
        {
            using (FileStream fs = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                workbook = new XSSFWorkbook(fs);
                fs.Close();

                Sheet = workbook.GetSheet(sheet);
            }
        }


        public int RowMax { get { return Sheet.LastRowNum; } }
        public string SheetName { get { return Sheet.SheetName; } }
        public XSSFRow CreateRow(int row, int cellCount = 0)
        {
            XSSFRow Row = Sheet.GetRow(row) as XSSFRow;
            if (Row == null)
                Row = Sheet.CreateRow(row) as XSSFRow;
            for (int i = 0; i < cellCount; i++)
            {
                Row.CreateCell(i);
            }
            return Row;
        }

        public XSSFCell CreateCell(int row, int column)
        {
            XSSFRow Row = Sheet.GetRow(row) as XSSFRow;
            if (Row == null)
                Row = Sheet.CreateRow(row) as XSSFRow;

            return Row.CreateCell(column) as XSSFCell;
        }

        public XSSFRow GetRow(int row) => Sheet.GetRow(row) as XSSFRow;
        public XSSFCell GetCell(int row, int column) => Sheet.GetRow(row).GetCell(column) as XSSFCell;
        public XSSFRow DeleteRow(int row) => Sheet.GetRow(row) as XSSFRow;

        /// <summary>
        /// 清除行数据
        /// </summary>
        /// <param name="startRow">从第几行开始</param>
        /// <param name="count">共清除count行</param>
        public void RemoveRow(int row, int count = 1)
        {
            for (int i = 0; i < count; i++)
            {
                Sheet.RemoveRow(GetRow(i));
            }
        }

        /// <summary>
        /// 删除行
        /// </summary>
        public void DeleteRow(int startRow, int endRow)
        {
            Sheet.ShiftRows(startRow, endRow, -1);
        }

        public void CreateValue(int row, int column, int value) => CreateCell(row, column).SetCellValue(value);
        public void CreateValue(int row, int column, string value) => CreateCell(row, column).SetCellValue(value);
        public void CreateValue(int row, int column, double value) => CreateCell(row, column).SetCellValue(value);
        public void CreateValue(int row, int column, bool value) => CreateCell(row, column).SetCellValue(value);
        public void CreateValue(int row, int column, DateTime value) => CreateCell(row, column).SetCellValue(value);

        public void UpdateValue(int row, int column, string value) => GetCell(row, column).SetCellValue(value);
        public void UpdateValue(int row, int column, int value) => GetCell(row, column).SetCellValue(value);
        public void UpdateValue(int row, int column, double value) => GetCell(row, column).SetCellValue(value);
        public void UpdateValue(int row, int column, DateTime value) => GetCell(row, column).SetCellValue(value);
        public void UpdateValue(int row, int column, bool value) => GetCell(row, column).SetCellValue(value);

        public string GetValue(int row, int column) => GetCell(row, column).ToString();
        public int GetValueInt(int row, int column) => (int)GetCell(row, column).NumericCellValue;
        public Double GetValueDouble(int row, int column) => GetCell(row, column).NumericCellValue;
        public String GetValueString(int row, int column) => GetCell(row, column).StringCellValue;
        public bool GetValueBool(int row, int column) => GetCell(row, column).BooleanCellValue;
        public DateTime GetValueDate(int row, int column) => GetCell(row, column).DateCellValue;

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


        public List<List<string>> Read(int row = 0, int column = 0)
        {
            try
            {
                List<List<string>> Values = new List<List<string>>();
                for (int i = 0; i <= Sheet.LastRowNum; i++)
                {
                    List<string> values = new List<string>();
                    XSSFRow Row = GetRow(i);
                    for (int j = 0; j < Row.LastCellNum; j++)
                    {
                        string value = GetValue(i, j);
                        values.Add(value);
                    }
                    Values.Add(values);
                }
                return Values;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }
        public List<List<int>> ReadInt(int row = 0, int column = 0)
        {
            try
            {
                List<List<int>> Values = new List<List<int>>();
                for (int i = 0; i <= Sheet.LastRowNum; i++)
                {
                    List<int> values = new List<int>();
                    XSSFRow Row = GetRow(i);
                    for (int j = 0; j < Row.LastCellNum; j++)
                    {
                        int value = GetValueInt(i, j);
                        values.Add(value);
                    }
                    Values.Add(values);
                }
                return Values;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }

        public List<List<double>> ReadDouble(int row = 0, int column = 0)
        {
            try
            {
                List<List<double>> Values = new List<List<double>>();
                for (int i = 0; i <= Sheet.LastRowNum; i++)
                {
                    List<double> values = new List<double>();
                    XSSFRow Row = GetRow(i);
                    for (int j = 0; j < Row.LastCellNum; j++)
                    {
                        double value = GetValueInt(i, j);
                        values.Add(value);
                    }
                    Values.Add(values);
                }
                return Values;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }

        public bool Write(List<List<string>> Values)
        {
            try
            {
                for (int i = 0; i < Values.Count; i++)
                {
                    List<string> values = Values[i];
                    for (int j = 0; j < values.Count; j++)
                    {
                        IRow row = Sheet.GetRow(i);
                        if (row == null)
                            row = Sheet.CreateRow(i);
                        ICell cell = row.GetCell(j);
                        if (cell == null)
                            cell = row.CreateCell(j);

                        cell.SetCellValue(values[j]);
                    }
                    Values.Add(values);
                }

                Write();
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }

        public bool Write(int row, int column, string vlaue)
        {
            try
            {
                IRow IRow = Sheet.GetRow(row);
                if (IRow == null)
                    IRow = Sheet.CreateRow(row);
                ICell cell = IRow.GetCell(column);
                if (cell == null)
                    cell = IRow.CreateCell(column);

                cell.SetCellValue(vlaue);
                Write();
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }


        public void Write()
        {
            using (FileStream fileStream = File.Open(path, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                workbook.Write(fileStream);
                fileStream.Close();
            
            }
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
    }
}
