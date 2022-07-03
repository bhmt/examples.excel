using System;
using System.Net.Mime;
using bhmt.examples.excel.core.Results;
using bhmt.examples.excel.core.Writer;
using bhmt.examples.excel.infrastructure.Settings;
using Microsoft.AspNetCore.Mvc;

namespace bhmt.examples.excel.api.Controllers
{
    /// <summary>
    /// Defines test controller for reporting results
    /// </summary>
    [ApiController]
    [Route("api/[controller]")]

    public class ExcelWriterController : ControllerBase
    {
        private readonly IWriterFactory _factory;

        public ExcelWriterController(IWriterFactory factory, FileStorageSettings settings)
        {
            _factory = factory ?? throw new ArgumentNullException(nameof(factory));
        }

        [HttpGet]
        [Route("template/{name}")]
        public IActionResult GetExcelFromTemplate([FromRoute] string name)
        {
            var w = _factory.GetWriter(true);
            var result = w.GetExcel(name);

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
            var w = _factory.GetWriter(false);
            var result = w.GetExcel(name);
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
