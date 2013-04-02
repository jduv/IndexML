﻿namespace IndexML.UnitTests.Wordprocessing
{
    using System;
    using System.Linq;
    using DocumentFormat.OpenXml.Wordprocessing;
    using IndexML.Wordprocessing;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    public class TableIndexerUnitTests : WordprocessingDocumentTest
    {
        #region Test Methods

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Constructor_NullArgument_ThrowsException()
        {
            var target = new TableIndexer(null);
        }

        [TestMethod]
        [DeploymentItem(StandardDocPath, TestFilesDir)]
        public void ImplicitCast_ValidIndexer_SameReference()
        {
            SafeExecuteTest(
               StandardDocPath,
               (doc) =>
               {
                   var expected = doc.MainDocumentPart.Document.Body.Descendants<Table>().FirstOrDefault();
                   var indexer = new TableIndexer(expected);
                   var target = (Table)indexer;

                   // Check references
                   Assert.AreSame(expected, target);
               });
        }

        [TestMethod]
        public void ImplicitCast_Null_IsNull()
        {
            TableIndexer indexer = null;
            var target = (Table)indexer;
            Assert.IsNull(target);
        }

        #endregion
    }
}
