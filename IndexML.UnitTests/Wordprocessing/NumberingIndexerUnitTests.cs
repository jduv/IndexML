﻿namespace IndexML.UnitTests.Wordprocessing
{
    using System;
    using DocumentFormat.OpenXml.Wordprocessing;
    using IndexML.Wordprocessing;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [DeploymentItem(@"TestFiles\", @"TestFiles\")]
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
