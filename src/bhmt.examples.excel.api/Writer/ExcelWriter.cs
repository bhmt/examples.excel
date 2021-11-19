using bhmt.examples.excel.api.Results;

namespace bhmt.examples.excel.api.Writer
{
    public class ExcelTemplateWriter : BaseExcelWriter, IExcelWriter
    {
        public readonly string sheetName = "Dummy";
        public ExcelTemplateWriter(FileStorageSettings settings) : base(settings) { }

        /// <inheritdoc />
        public MemoryStreamResult GetExcel(string name)
        {
            SetSheet(sheetName);

            var valCol = 2;
            var date = Utils.FormatNow();

            SetCell(ExcelWriterConfig.NameRow, valCol, name);
            SetCell(ExcelWriterConfig.DateRow, valCol, date);
            SetCell(ExcelWriterConfig.RandRow, valCol, Utils.RandStr(15));

            return SaveMemoryStream($"ExcelFromTemplate_{name}_{date}.xlsx".Replace(" ", string.Empty));
        }
    }

    public class ExcelCodeWriter : BaseExcelWriter, IExcelWriter
    {
        public readonly string sheetName = "Dummy";
        public ExcelCodeWriter() : base(null) { }

        /// <inheritdoc />
        public MemoryStreamResult GetExcel(string name)
        {
            SetSheet(sheetName);

            int keyCol = 1, valCol = 2;
            var date = Utils.FormatNow();


            SetCell(ExcelWriterConfig.NameRow, keyCol, "Name");
            SetCell(ExcelWriterConfig.NameRow, valCol, name);

            SetCell(ExcelWriterConfig.DateRow, keyCol, "Date");
            SetCell(ExcelWriterConfig.DateRow, valCol, date);

            SetCell(ExcelWriterConfig.RandRow, keyCol, "Rand");
            SetCell(ExcelWriterConfig.RandRow, valCol, Utils.RandStr(15));

            return SaveMemoryStream($"ExcelFromCode_{name}_{date}.xlsx".Replace(" ", string.Empty));
        }
    }
}
