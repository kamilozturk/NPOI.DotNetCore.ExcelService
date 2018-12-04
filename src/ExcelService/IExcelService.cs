using System.Collections.Generic;
using System.IO;

namespace NPOI.DotNetCore.ExcelService
{
    public interface IExcelService
    {
        string DefaultContentType { get; }
        byte[] ExportAsExcel(bool useXLSXFormat, string sheetName, IEnumerable<string> header, IEnumerable<IEnumerable<object>> data);
        IEnumerable<IEnumerable<object>> ImportExcel(bool isXLSXFormat, Stream stream, bool skipFistRow);
    }
}
