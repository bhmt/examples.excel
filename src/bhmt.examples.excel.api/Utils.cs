using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace bhmt.examples.excel.api
{
    public static class Utils
    {
        private const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";

        /// <summary>
        /// Return a random alphanumeric string with the given length.
        /// </summary>
        /// <param name="length">Positive integer.</param>
        /// <returns></returns>
        public static string RandStr(int length)
        {
            Random random = new();
            return length < 0
                ? throw new ArgumentException("The length of the string must be positive.")
                : new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }

        public static string FormatNow() => DateTime.Now.ToString("yyyy-MM-dd");
    }
}
