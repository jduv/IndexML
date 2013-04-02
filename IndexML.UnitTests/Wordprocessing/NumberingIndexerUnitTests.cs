namespace IndexML.UnitTests.Wordprocessing
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
    using IndexML.Wordprocessing;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    public class NumberingIndexerUnitTests : WordprocessingDocumentTest
    {
        #region Test Methods

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Constructor_NullArgument_ThrowsException()
        {
            var target = new NumberingIndexer(null);
        }

        [TestMethod]
        [DeploymentItem(NumberedDocPath, TestFilesDir)]
        public void ImplicitCast_ValidIndexer_SameReference()
        {
            SafeExecuteTest(
               NumberedDocPath,
               (doc) =>
               {
                   var expected = doc.MainDocumentPart.NumberingDefinitionsPart.Numbering;
                   var indexer = new NumberingIndexer(doc.MainDocumentPart.NumberingDefinitionsPart);
                   var target = (Numbering)indexer;

                   // Check references
                   Assert.AreSame(expected, target);
               });
        }

        [TestMethod]
        public void ImplicitCast_Null_IsNull()
        {
            NumberingIndexer indexer = null;
            var target = (Numbering)indexer;
            Assert.IsNull(target);
        }

        #endregion
    }
}
