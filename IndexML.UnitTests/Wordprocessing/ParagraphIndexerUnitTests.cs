namespace IndexML.UnitTests.Wordprocessing
{
    using System;
    using System.Linq;
    using DocumentFormat.OpenXml.Wordprocessing;
    using IndexML.Wordprocessing;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    public class ParagraphIndexerUnitTests : WordprocessingDocumentTest
    {
        #region Test Methods

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Constructor_NullArgument_ThrowsException()
        {
            var target = new ParagraphIndexer(null);
        }

        [TestMethod]
        [DeploymentItem(StandardDocPath, TestFilesDir)]
        public void Constructor_StandardDoc_ValidElements()
        {
            SafeExecuteTest(
                StandardDocPath,
                (doc) =>
                {
                    var expected = doc.MainDocumentPart.Document.Body.Elements<Paragraph>().FirstOrDefault();
                    var target = new ParagraphIndexer(expected);

                    // Check to see if references match up and are valid.
                    Assert.IsNotNull(target.Paragraph);
                    Assert.AreSame(expected, target.Paragraph);
                    Assert.IsNotNull(target.Properties);
                    Assert.AreSame(expected.Elements<ParagraphProperties>().FirstOrDefault(), target.Properties);
                });
        }

        #endregion
    }
}
