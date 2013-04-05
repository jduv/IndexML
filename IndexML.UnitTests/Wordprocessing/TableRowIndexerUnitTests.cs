namespace IndexML.UnitTests.Wordprocessing
{
    using System;
    using DocumentFormat.OpenXml.Wordprocessing;
    using IndexML.Wordprocessing;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    
    [TestClass]
    public class TableRowIndexerUnitTests : WordprocessingDocumentTest
    {
        #region Test Methods

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Constructor_NullArgument_ThrowsException()
        {
            var target = new TableRowIndexer(null);
        }

        [TestMethod]
        public void ImplicitCast_ValidIndexer_SameReference()
        {
            var expected = new TableRow();
            var indexer = new TableRowIndexer(expected);
            var target = (TableRow)indexer;

            Assert.AreSame(expected, indexer.Row);
        }

        [TestMethod]
        public void ImplicitCast_Null_IsNull()
        {
            TableRowIndexer indexer = null;
            var target = (TableRow)indexer;
            Assert.IsNull(target);
        }

        #endregion
    }
}
