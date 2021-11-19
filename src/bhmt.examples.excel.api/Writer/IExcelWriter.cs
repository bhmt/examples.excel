using bhmt.examples.excel.api.Results;

namespace bhmt.examples.excel.api.Writer
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
