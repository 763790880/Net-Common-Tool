using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassCommon
{
    public class NPOIExcleCommon
    {
        /// <summary>
        /// DataTable转换成Excel文档流(导出数据量超出65535条,分sheet)
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        public static void ExportDataTableToExcel(DataTable sourceTable, string filePath)
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            MemoryStream ms = new MemoryStream();
            int dtRowsCount = sourceTable.Rows.Count;
            int SheetCount = Convert.ToInt32(Math.Ceiling(Convert.ToDouble(dtRowsCount) / 65536));
            int SheetNum = 1;
            int rowIndex = 1;
            int tempIndex = 1; //标示 
            ISheet sheet = workbook.CreateSheet("sheet1" + SheetNum);
            ICellStyle cellStyle = workbook.CreateCellStyle();//创建样式对象
            cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            var fB = workbook.CreateFont();
            fB.Boldweight = (short)FontBoldWeight.Bold;
            var fN = workbook.CreateFont();
            fN.Boldweight = (short)FontBoldWeight.None;
            try
            {
                for (int i = 0; i < dtRowsCount; i++)
                {
                    if (i == 0 || tempIndex == 1)
                    {
                        IRow headerRow = sheet.CreateRow(0);
                        foreach (DataColumn column in sourceTable.Columns)
                        {
                            var cell = headerRow.CreateCell(column.Ordinal);
                            cell.SetCellValue(column.ColumnName);
                            var f = workbook.CreateFont();
                            f.Boldweight = (short)FontBoldWeight.Bold;
                            ICellStyle style = workbook.CreateCellStyle();//创建样式对象
                            style.Alignment = HorizontalAlignment.Center;
                            style.FillForegroundColor = 55;
                            //style.FillPattern = FillPatternType.SOLID_FOREGROUND;
                            style.FillPattern = FillPattern.SolidForeground;
                            style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                            style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                            style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                            style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                            style.SetFont(f);
                            cell.CellStyle = style;
                        }
                    }
                    HSSFRow dataRow = (HSSFRow)sheet.CreateRow(tempIndex);
                    foreach (DataColumn column in sourceTable.Columns)
                    {
                        var cell = dataRow.CreateCell(column.Ordinal);
                        var val = GetValue(sourceTable.Rows[i][column]);
                        cell.SetCellValue(val);
                        cell.CellStyle = cellStyle;
                    }
                    if (tempIndex == 65535)
                    {
                        SheetNum++;
                        sheet = workbook.CreateSheet("sheet" + SheetNum);//
                        tempIndex = 0;
                    }
                    rowIndex++;
                    tempIndex++;
                    //AutoSizeColumns(sheet);
                }
            }
            catch (Exception ex)
            {

                throw;
            }


            workbook.Write(ms);
            ms.Flush();
            ms.Position = 0;
            sheet = null;
            // headerRow = null;
            workbook = null;


            var buf = ms.ToArray();

            //保存为Excel文件  
            using (FileStream fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                fs.Write(buf, 0, buf.Length);
                fs.Flush();
            }
        }


        /// <summary>
        /// List转换成Excel文档流(导出数据量超出65535条,分sheet)
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        public static void ExportListToExcel<T>(List<T> sourceTable, string filePath)
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            MemoryStream ms = new MemoryStream();
            int dtRowsCount = sourceTable.Count;
            int SheetCount = Convert.ToInt32(Math.Ceiling(Convert.ToDouble(dtRowsCount) / 65536));
            int SheetNum = 1;
            int rowIndex = 1;
            int tempIndex = 1; //标示 
            ISheet sheet = workbook.CreateSheet("sheet1" + SheetNum);
            for (int i = 0; i < dtRowsCount; i++)
            {
                Type t = typeof(T);
                var Properties = t.GetProperties();
                if (i == 0 || tempIndex == 1)
                {
                    IRow headerRow = sheet.CreateRow(0);
                    for (int j = 0; j < Properties.Length; j++)
                    {
                        var name = "";
                        object[] objs = Properties[j].GetCustomAttributes(typeof(DescriptionAttribute), false);    //获取描述属性
                        if (objs == null || objs.Length == 0)    //当描述属性没有时，直接返回名称
                            name = Properties[j].Name;
                        else
                        {
                            DescriptionAttribute descriptionAttribute = (DescriptionAttribute)objs[0];
                            name = descriptionAttribute.Description;
                        }

                        headerRow.CreateCell(j).SetCellValue(name);
                    }

                }
                HSSFRow dataRow = (HSSFRow)sheet.CreateRow(tempIndex);
                for (int l = 0; l < Properties.Length; l++)
                {
                    var val = Properties[l].GetValue(t);
                    dataRow.CreateCell(l).SetCellValue(val.ToString());
                }
                if (tempIndex == 65535)
                {
                    SheetNum++;
                    sheet = workbook.CreateSheet("sheet" + SheetNum);//
                    tempIndex = 0;
                }
                rowIndex++;
                tempIndex++;
                //AutoSizeColumns(sheet);
            }

            workbook.Write(ms);
            ms.Flush();
            ms.Position = 0;
            sheet = null;
            // headerRow = null;
            workbook = null;


            var buf = ms.ToArray();

            //保存为Excel文件  
            using (FileStream fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                fs.Write(buf, 0, buf.Length);
                fs.Flush();
            }
        }
        /// <summary>
        /// Excel导入成Datable
        /// </summary>
        /// <param name="file">导入路径(包含文件名与扩展名)</param>
        /// <returns></returns>
        public static DataTable ExcelToTable(string file)
        {
            DataTable dt = new DataTable();
            IWorkbook workbook;
            string fileExt = Path.GetExtension(file).ToLower();
            using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
            {
                //XSSFWorkbook 适用XLSX格式，HSSFWorkbook 适用XLS格式
                if (fileExt == ".xlsx") { workbook = new XSSFWorkbook(fs); } else if (fileExt == ".xls") { workbook = new HSSFWorkbook(fs); } else { workbook = null; }
                if (workbook == null) { return null; }
                ISheet sheet = workbook.GetSheetAt(0);

                //表头  
                IRow header = sheet.GetRow(sheet.FirstRowNum);
                List<int> columns = new List<int>();
                for (int i = 0; i < header.LastCellNum; i++)
                {
                    object obj = GetValueType(header.GetCell(i));
                    if (obj == null || obj.ToString() == string.Empty)
                    {
                        dt.Columns.Add(new DataColumn("Columns" + i.ToString()));
                    }
                    else
                        dt.Columns.Add(new DataColumn(obj.ToString()));
                    columns.Add(i);
                }
                //数据  
                for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                {
                    DataRow dr = dt.NewRow();
                    bool hasValue = false;
                    if (sheet.GetRow(i) != null)//当前行不为空
                    {
                        foreach (int j in columns)
                        {
                            dr[j] = GetValueType(sheet.GetRow(i).GetCell(j));
                            if (dr[j] != null && dr[j].ToString() != string.Empty)
                            {
                                hasValue = true;
                            }
                        }
                    }
                    if (hasValue)
                    {
                        dt.Rows.Add(dr);
                    }
                }
            }
            return dt;
        }
        /// <summary>
        /// Excel导入成Datable
        /// </summary>
        /// <param name="extend">扩展名</param>
        ///  <param name="fileStream">对应流</param>
        /// <returns></returns>
        public static DataTable ExcelToTable(string extend, Stream fileStream)
        {
            try
            {
                DataTable dt = new DataTable();
                IWorkbook workbook;
                //XSSFWorkbook 适用XLSX格式，HSSFWorkbook 适用XLS格式
                if (extend == ".xlsx") { workbook = new XSSFWorkbook(fileStream); } else if (extend == ".xls") { workbook = new XSSFWorkbook(fileStream); } else { workbook = null; }
                if (workbook == null) { return null; }
                ISheet sheet = workbook.GetSheetAt(0);

                //表头  
                IRow header = sheet.GetRow(sheet.FirstRowNum);
                List<int> columns = new List<int>();
                for (int i = 0; i < header.LastCellNum; i++)
                {
                    object obj = GetValueType(header.GetCell(i));
                    if (obj == null || obj.ToString() == string.Empty)
                    {
                        dt.Columns.Add(new DataColumn("Columns" + i.ToString()));
                    }
                    else
                        dt.Columns.Add(new DataColumn(obj.ToString()));
                    columns.Add(i);
                }
                //数据  
                for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                {
                    DataRow dr = dt.NewRow();
                    bool hasValue = false;
                    foreach (int j in columns)
                    {
                        dr[j] = GetValueType(sheet.GetRow(i).GetCell(j));
                        if (dr[j] != null && dr[j].ToString() != string.Empty)
                        {
                            hasValue = true;
                        }
                    }
                    if (hasValue)
                    {
                        dt.Rows.Add(dr);
                    }
                }
                //fileStream.Close();
                return dt;
            }
            catch (Exception ex)
            {

                throw;
            }

        }

        /// <summary>
        /// 获取单元格类型
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private static object GetValueType(ICell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                case CellType.Blank: //BLANK:  
                    return null;
                case CellType.Boolean: //BOOLEAN:  
                    return cell.BooleanCellValue;
                case CellType.Numeric: //NUMERIC:  
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        return cell.DateCellValue.ToString("yyyy/MM/dd HH:mm:ss");
                    }
                    return cell.NumericCellValue;
                case CellType.String: //STRING:  
                    return cell.StringCellValue;
                case CellType.Error: //ERROR:  
                    return cell.ErrorCellValue;
                case CellType.Formula: //FORMULA:  
                    cell.SetCellType(CellType.String);
                    return cell.StringCellValue;
                default:
                    return "=" + cell.CellFormula;
            }
        }
        /// <summary>
        /// 对应值转换为EXCEL识别的值
        /// </summary>
        /// <param name="val"></param>
        /// <returns></returns>
        private static dynamic GetValue(object val)
        {
            dynamic v;
            var b = double.TryParse(val.ToString(), out double dl);
            b = int.TryParse(val.ToString(), out int i);
            b = decimal.TryParse(val.ToString(), out decimal d);
            b = float.TryParse(val.ToString(), out float f);
            if (b)
                return Convert.ToDouble(val);
            b = bool.TryParse(val.ToString(), out bool bl);
            if (b)
                return Convert.ToBoolean(val);
            b = DateTime.TryParse(val.ToString(), out DateTime dt);
            if (b)
                return Convert.ToDateTime(val);
            return val.ToString();
        }
    }
}
