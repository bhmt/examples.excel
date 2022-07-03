using System;
using bhmt.examples.excel.core.Writer;
using bhmt.examples.excel.infrastructure.Settings;


namespace bhmt.examples.excel.infrastructure.Writer
{
    public class WriterFactory : IWriterFactory
    {
        private readonly FileStorageSettings _settings;
        public WriterFactory(FileStorageSettings settings)
        {
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
        }

        /// <summary>
        /// Get theproper writer implementation.
        /// </summary>
        public IExcelWriter GetWriter(bool isTemplate) => isTemplate ? new ExcelTemplateWriter(_settings) : new ExcelCodeWriter();
    }
}
