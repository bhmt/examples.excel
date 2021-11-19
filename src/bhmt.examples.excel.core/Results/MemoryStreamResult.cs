namespace bhmt.examples.excel.core.Results
{
    public class MemoryStreamResult
    {
        public byte[] Buffer { get; set; }
        public string Name { get; set; }

        public MemoryStreamResult(byte[] buffer, string name)
        {
            Buffer = buffer;
            Name = name;
        }
    }
}
