namespace IndexML.UnitTests.Wordprocessing
{
    using System;
    using System.IO;
    using DocumentFormat.OpenXml.Packaging;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A base test class for testing document indexers.
    /// </summary>
    public abstract class WordprocessingDocumentTest : Test
    {
        #region Fields & Constants

        protected const string EmptyDocPath = TestFilesDir + "Empty.docx";

        #endregion

        #region Protected Methods

        protected static WordprocessingDocument LoadTestDoc(string path)
        {
            AssertFileExists(path);

            try
            {
                using (var fileStream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    var memory = new MemoryStream(); // Make the stream expandable by using default ctor
                    CopyStream(fileStream, memory);      // Copy the stream to memory so we can do whatever we want with it
                    return WordprocessingDocument.Open(memory, true);
                }
            }
            catch (Exception exc)
            {
                Assert.Inconclusive("Test inconclusive. Unable to open the document at path " + path + ". Exception: " + exc.Message);
            }

            return null;
        }

        protected static void SafeExecuteTest<TActionable>(
            string docPath,
            Func<WordprocessingDocument, TActionable> selector,
            Action<TActionable> testToPerform)
        {
            if (testToPerform == null)
            {
                Assert.Inconclusive("No test specified to perform!");
            }

            var doc = LoadTestDoc(docPath);
            if (doc != null)
            {
                using (doc)
                {
                    var items = selector == null ? default(TActionable) : selector(doc);
                    testToPerform(items);
                }
            }
        }

        protected static void SafeExecuteTest(
            string docPath,
            Action<WordprocessingDocument> testToPerform)
        {
            if (testToPerform == null)
            {
                Assert.Inconclusive("No test specified to perform!");
            }

            var doc = LoadTestDoc(docPath);
            if (doc != null)
            {
                using (doc)
                {
                    testToPerform(doc);
                }
            }
        }

        #endregion
    }
}
