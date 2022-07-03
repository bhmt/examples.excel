namespace bhmt.examples.excel.core.Writer
{
    public interface IWriterFactory
    {
        /// <summary>
        /// Get the writer implementation.
        /// </summary>
        IExcelWriter GetWriter(bool isTemplate);
    }
}
