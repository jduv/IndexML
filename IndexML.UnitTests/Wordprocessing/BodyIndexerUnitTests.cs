namespace IndexML.UnitTests.Wordprocessing
{
    using System;
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
    using IndexML.Wordprocessing;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    public class BodyIndexerUnitTests : WordprocessingDocumentTest
    {
        #region Test Methods

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Constructor_NullArgument_ThrowsException()
        {
            var target = new BodyIndexer(null);
        }

        [TestMethod]
        [DeploymentItem(EmptyDocPath, TestFilesDir)]
        public void Constructor_EmptyDocument_ValidState()
        {
            SafeExecuteTest(
                EmptyDocPath,
                (doc) =>
                {
                    var target = new BodyIndexer(doc.MainDocumentPart.Document.Body);

                    Assert.IsNotNull(target.Paragraphs);
                    Assert.AreEqual(1, target.Paragraphs.Count()); // Only one paragraph in an empty doc.
                    Assert.IsNotNull(target.Tables);
                    Assert.AreEqual(0, target.Tables.Count()); // No tables.
                    Assert.IsNotNull(target.FinalSectionProperties);

                    // Check references.
                    Assert.AreSame(
                        doc.MainDocumentPart.Document.Body.Elements<SectionProperties>().FirstOrDefault(), 
                        target.FinalSectionProperties);
                });
        }

        [TestMethod]
        [DeploymentItem(StandardDocPath, TestFilesDir)]
        public void Constructor_StandardDocument_ValidState()
        {
            SafeExecuteTest(
                StandardDocPath,
                (doc) =>
                {
                    var target = new BodyIndexer(doc.MainDocumentPart.Document.Body);

                    Assert.IsNotNull(target.Paragraphs);
                    Assert.IsTrue(target.Paragraphs.Count() > 1);
                    Assert.IsNotNull(target.Tables);
                    Assert.IsTrue(target.Tables.Count() > 0);
                    Assert.IsNotNull(target.FinalSectionProperties);

                    // Check references.
                    Assert.AreSame(
                        doc.MainDocumentPart.Document.Body.Elements<SectionProperties>().FirstOrDefault(),
                        target.FinalSectionProperties);
                });
        }

        [TestMethod]
        [DeploymentItem(StandardDocPath, TestFilesDir)]
        public void ImplicitCast_ValidIndexerSameReference()
        {
            AssertFileExists(StandardDocPath);
            using (var target = new WordprocessingDocumentIndexer(OpenFileReadWrite(StandardDocPath)))
            {
                WordprocessingDocument doc = (WordprocessingDocument)target;
                Assert.IsNotNull(doc);
                Assert.AreSame(target.WordprocessingDocument, doc);
            }
        }

        #endregion
    }
}
