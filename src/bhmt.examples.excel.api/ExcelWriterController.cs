using System;
using System.Net.Mime;

using Microsoft.AspNetCore.Mvc;
using bhmt.examples.excel.core.Results;
using bhmt.examples.excel.infrastructure.Writer;

namespace bhmt.examples.excel.api.Controllers
{
    /// <summary>
    /// Defines test controller for reporting results
    /// </summary>
    [ApiController]
    [Route("api/[controller]")]

    public class ExcelWriterController : ControllerBase
    {
        private readonly IExcelWriter excelTemplateWriter;
        private readonly IExcelWriter excelCodeWriter;

        public ExcelWriterController(IExcelWriter excelTemplateWriter, IExcelWriter excelCodeWriter)
        {
            this.excelTemplateWriter = excelTemplateWriter ?? throw new ArgumentNullException(nameof(excelTemplateWriter));
            this.excelCodeWriter = excelCodeWriter ?? throw new ArgumentNullException(nameof(excelCodeWriter));
        }

        [HttpGet]
        [Route("template/{name}")]
        public IActionResult GetExcelFromTemplate([FromRoute] string name)
        {
            var result = excelTemplateWriter.GetExcel(name);

            Response.Headers["Content-Disposition"] = new ContentDisposition
            {
                FileName = result.Name,
                Inline = false
            }.ToString();

            return File(result.Buffer, "application/octet-stream");
        }

        [HttpGet]
        [Route("code/{name}")]
        public IActionResult GetExcelFromCode([FromRoute] string name)
        {
            var result = excelCodeWriter.GetExcel(name);
            return ToFile(result);
        }

        private FileContentResult ToFile(MemoryStreamResult result)
        {
            Response.Headers["Content-Disposition"] = new ContentDisposition
            {
                FileName = result.Name,
                Inline = false
            }.ToString();

            return File(result.Buffer, "application/octet-stream");
        }
    }
}
