using System.IO;

namespace DocumentFormat.OpenXml.Extensions
{
    public class AbstractWriter
    {
        ///<summary>
        /// Writes the contents of a stream to a file.
        /// </summary>
        public static void StreamToFile(string path, MemoryStream stream)
        {
            using (var file = new FileStream(path, FileMode.Create))
            {
                stream.WriteTo(file);
                file.Close();
            }
        }
    }
}
