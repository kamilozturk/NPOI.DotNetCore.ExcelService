using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace NPOI.DotNetCore.ExcelService.DefaultImpl
{
    public class ExcelService : IExcelService
    {
        public string DefaultContentType => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        public byte[] ExportAsExcel(bool useXLSXFormat, string sheetName, IEnumerable<string> header, IEnumerable<IEnumerable<object>> data)
        {
            using (var memory = new MemoryStream())
            {
                IWorkbook workbook = useXLSXFormat ? (IWorkbook)new XSSFWorkbook() : new HSSFWorkbook();
                ISheet excelSheet = workbook.CreateSheet(sheetName);

                var nextRowNumber = WriteRow(header, excelSheet, 0);

                foreach (var item in data)
                    nextRowNumber = WriteRow(item, excelSheet, nextRowNumber);

                workbook.Write(memory);

                return memory.ToArray();
            }
        }

        private int WriteRow(IEnumerable<object> data, ISheet excelSheet, int rowNumber)
        {
            if (data == null)
                return rowNumber;

            IRow row = excelSheet.CreateRow(rowNumber);

            var cellNumber = 0;

            foreach (var item in data)
            {
                var type = item.GetType();
                var cell = row.CreateCell(cellNumber++);

                if (type == typeof(double))
                    cell.SetCellValue((double)item);
                else if (type == typeof(bool))
                    cell.SetCellValue((bool)item);
                else if (type == typeof(DateTime))
                    cell.SetCellValue((DateTime)item);
                else
                    cell.SetCellValue(item.ToString());
            }

            return rowNumber + 1;
        }

        public IEnumerable<IEnumerable<object>> ImportExcel(bool isXLSXFormat, Stream stream, bool skipFistRow)
        {
            ISheet sheet;

            if (!isXLSXFormat)
            {
                HSSFWorkbook hssfwb = new HSSFWorkbook(stream); //This will read the Excel 97-2000 formats  
                sheet = hssfwb.GetSheetAt(0); //get first sheet from workbook  
            }
            else
            {
                XSSFWorkbook hssfwb = new XSSFWorkbook(stream); //This will read 2007 Excel format  
                sheet = hssfwb.GetSheetAt(0); //get first sheet from workbook   
            }

            var rowIndex = skipFistRow ? 1 : 0;
            var firstRow = sheet.GetRow(rowIndex); //Get Header Row
            var cellCount = firstRow.LastCellNum;
            var data = new List<List<string>>();

            for (int i = rowIndex; i <= sheet.LastRowNum; i++) //Read Excel File
            {
                var row = sheet.GetRow(i);

                if (row == null) continue;
                if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;

                var rowData = new List<string>();

                for (int j = row.FirstCellNum; j < cellCount; j++)
                {
                    var value = row.GetCell(j).ToString();

                    rowData.Add(value);
                }

                data.Add(rowData);
            }

            return data;
        }
    }
}
