namespace IndexML.UnitTests.Wordprocessing
{
    using System;
    using IndexML.Wordprocessing;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    class DocumentIndexerUnitTests : WordprocessingDocumentTest
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
                });
        }

        #endregion
    }
}
