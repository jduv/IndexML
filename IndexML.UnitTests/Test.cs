namespace IndexML.UnitTests
{
    using System.IO;

    /// <summary>
    /// A base test class. Defines some re-usable logic.
    /// </summary>
    public class Test
    {
        #region Fields & Constants

        protected const string TestFilesDir = @"TestFiles\";

        #endregion

        #region Protected Methods

        protected static byte[] ReadAllBytes(string path)
        {
            // Helper method, assumes the file exists.
            var memory = new MemoryStream();
            using (var fileStream = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
            {
                CopyStream(fileStream, memory);
            }

            return memory.ToArray();
        }

        protected static void CopyStream(Stream source, Stream target)
        {
            // Helper method, assumes source and target are not null.
            using (source)
            {
                var buffer = new byte[32768];
                int bytesRead;
                while ((bytesRead = source.Read(buffer, 0, buffer.Length)) > 0)
                {
                    target.Write(buffer, 0, bytesRead);
                }
            }
        }

        #endregion
    }
}
