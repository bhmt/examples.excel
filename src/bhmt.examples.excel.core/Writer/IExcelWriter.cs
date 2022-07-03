using bhmt.examples.excel.core.Results;

namespace bhmt.examples.excel.core.Writer
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
