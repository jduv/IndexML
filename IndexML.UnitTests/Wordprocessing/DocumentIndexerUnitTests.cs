namespace IndexML.UnitTests.Wordprocessing
{
    using System;
    using DocumentFormat.OpenXml.Wordprocessing;
    using IndexML.Wordprocessing;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    public class DocumentIndexerUnitTests : WordprocessingDocumentTest
    {
        #region Test Methods

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Constructor_NullArgument_ThrowsException()
        {
            var target = new DocumentIndexer(null);
        }

        [TestMethod]
        [DeploymentItem(EmptyDocPath, TestFilesDir)]
        public void Constructor_EmptyDocument_ValidState()
        {
            SafeExecuteTest(
                EmptyDocPath,
                (doc) => 
                {
                    var target = new DocumentIndexer(doc.MainDocumentPart);
                    Assert.IsNotNull(target);
                    Assert.IsNotNull(target.Document);
                    Assert.IsNotNull(target.Body);

                    // Check references
                    Assert.AreSame(doc.MainDocumentPart.Document, target.Document);
                });
        }

        [TestMethod]
        public void ImplicitCast_ValidIndexer_SameReference()
        {
            SafeExecuteTest(
               EmptyDocPath,
               (doc) =>
               {
                   var expected = doc.MainDocumentPart.Document;
                   var indexer = new DocumentIndexer(doc.MainDocumentPart);
                   var target = (Document)indexer;

                   // Check references
                   Assert.AreSame(expected, target);
               });
        }

        [TestMethod]
        public void ImplicitCast_Null_IsNull()
        {
            DocumentIndexer indexer = null;
            var target = (Document)indexer;
            Assert.IsNull(target);
        }

        #endregion
    }
}
