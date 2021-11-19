using System;
using System.IO;
using NetEscapades.Configuration.Validation;

namespace bhmt.examples.excel.api
{
    /// <summary>
    /// Defines validation class for file storage settings.
    /// </summary>
    public class FileStorageSettings : IValidatable
    {
        public string TemplatesPath { get; set; }
        public string TemplatesFolder { get; set; }
        public string ReportTemplateFile { get; set; }

        public void Validate()
        {
            if (TemplatesPath == null)
            {
                throw new ArgumentException("Templates folder path is required.", nameof(TemplatesPath));
            }

            if (TemplatesFolder == null)
            {
                throw new ArgumentException("Templates folder name is required.", nameof(TemplatesFolder));
            }

            if (ReportTemplateFile == null)
            {
                throw new ArgumentException("Reports template file name is required.", nameof(ReportTemplateFile));
            }
        }

        public string ReportTemplate() => Path.Combine(TemplatesPath, TemplatesFolder, ReportTemplateFile);
    }
}
