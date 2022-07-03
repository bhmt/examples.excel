using System.IO;
using bhmt.examples.excel.core.Results;
using bhmt.examples.excel.infrastructure.Settings;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace bhmt.examples.excel.infrastructure.Writer
{
    public class BaseExcelWriter
    {
        private readonly IWorkbook wb;
        private ISheet ws;
        // private string outputPath;
        private readonly XSSFFormulaEvaluator eval;

        /// <summary>
        /// Open an excel file from a the excel template.
        /// Create a workbook and an evaluator for formulas.
        /// </summary>
        /// <param name="settings"><see cref="FileStorageSettings"></param>
        public BaseExcelWriter(FileStorageSettings settings)
        {
            if (settings is null)
            {
                wb = new XSSFWorkbook();
                ws = wb.CreateSheet("Dummy");
            }
            else
            {
                using (var fs = new FileStream(settings.ReportTemplate(), FileMode.Open, FileAccess.Read))
                {
                    wb = WorkbookFactory.Create(settings.ReportTemplate());
                }
            }

            eval = new XSSFFormulaEvaluator(wb);
        }

        /// <summary>
        /// Return the instance of the class so it can be used in other writer implementations.
        /// Important to point to the same opened file.
        /// </summary>
        /// <returns></returns>
        protected BaseExcelWriter GetBase() => this;



        /// <summary>
        /// Set a sheet to write to using a name.
        /// </summary>
        /// <param name="sheetName"></param>
        public void SetSheet(string sheetName) => ws = wb.GetSheet(sheetName);

        /// <summary>
        /// Set a sheet to write to using an index.
        /// </summary>
        /// <param name="sheetIndex"></param>
        public void SetSheet(int sheetIndex) => ws = wb.GetSheetAt(sheetIndex);

        /// <summary>
        /// Check if the row and clumun indexes are bigger than 0
        /// </summary>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <returns></returns>
        private static bool CheckIndexes(int row, int column) => row > 0 && column > 0;

        /// <summary>
        /// Set a string value to a cell.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <param name="value"></param>
        public void SetCell(int row, int column, string value)
        {
            if (CheckIndexes(row, column))
            {
                var r = ws.GetRow(row - 1);
                if (r is null)
                {
                    r = ws.CreateRow(row - 1);
                }

                var c = r.GetCell(column - 1);
                if (c is null)
                {
                    r.CreateCell(column - 1);
                }

                ws.GetRow(row - 1).GetCell(column - 1).SetCellValue(value);
            }
        }

        /// <summary>
        /// Set an int value to a cell.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <param name="value"></param>
        public void SetCell(int row, int column, int value)
        {
            if (CheckIndexes(row, column))
            {
                var r = ws.GetRow(row - 1);
                if (r is null)
                {
                    r = ws.CreateRow(row - 1);
                }

                var c = r.GetCell(column - 1);
                if (c is null)
                {
                    r.CreateCell(column - 1);
                }

                ws.GetRow(row - 1).GetCell(column - 1).SetCellValue(value);
            }
        }

        /// <summary>
        /// Save data to a memory stream.
        /// </summary>
        public MemoryStreamResult SaveMemoryStream(string outputPath)
        {
            eval.EvaluateAll();
            MemoryStreamResult result;
            using (var ms = new MemoryStream())
            {
                wb.Write(ms);
                result = new MemoryStreamResult(ms.ToArray(), Path.GetFileName(outputPath));
            }

            return result;
        }

        ///// <summary>
        ///// Save and store data to excel file on disc.
        ///// </summary>
        //public void SaveFileStream(string outputPath)
        //{
        //    this.outputPath = outputPath;
        //    eval.EvaluateAll();
        //    using (var fs = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
        //    {
        //        wb.Write(fs);
        //    }
        //}

        // /// <summary>
        // /// Delete the file created by the writer.
        // /// </summary>s
        // public void Delete()
        // {
        //     if (File.Exists(outputPath))
        //     {
        //         File.Delete(outputPath);
        //     }
        // }
    }
}
