namespace IndexML.UnitTests.Wordprocessing
{
    using System;
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
        public void ImplicitCast_Null_IsNull()
        {
            TableIndexer indexer = null;
            var target = (Table)indexer;
            Assert.IsNull(target);
        }

        #endregion
    }
}
