namespace IndexML.UnitTests
{
    using System;
    using System.IO;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A base test class. Defines some re-usable logic.
    /// </summary>
    public class Test
    {
        #region Fields & Constants

        protected const string TestFilesDir = @"IndexML.TestFiles\";

        #endregion

        #region Protected Methods

        protected static void AssertFileExists(string path)
        {
            if (!File.Exists(path))
            {
                Assert.Inconclusive("Test inconclusive. A required file was not found! Path: " + path);
            }
        }

        protected static byte[] LoadTestFileBytes(string path)
        {
            AssertFileExists(path);

            try
            {
                return ReadAllBytes(path);
            }
            catch (Exception exc)
            {
                Assert.Inconclusive("Test inconclusive. Unable to read the document at path " + path + ". Exception: " + exc.Message);
            }

            return null;
        }

        protected static byte[] ReadAllBytes(string path)
        {
            // Helper method, assumes the file exists.
            var memory = new MemoryStream();
            using (var fileStream = OpenFileReadWrite(path))
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

        protected static FileStream OpenFileReadWrite(string path)
        {
            // Opens a file in Read/Write state.
            return File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
        }

        #endregion
    }
}
