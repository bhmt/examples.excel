using bhmt.examples.excel.core.Results;

namespace bhmt.examples.excel.infrastructure.Writer
{
    public interface IExcelWriter
    {
        /// <summary>
        /// Fill a cell and write the excel to a memory stream.
        /// Intended for creating a temporary downloadable file.
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        MemoryStreamResult GetExcel(string name);
    }
}
