namespace IndexML.UnitTests.Wordprocessing
{
    using System;
    using DocumentFormat.OpenXml.Wordprocessing;
    using IndexML.Wordprocessing;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    public class TableCellIndexerUnitTests
    {
        #region Test Methods

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Constructor_NullArgument_ThrowsException()
        {
            var target = new TableCellIndexer(null);
        }

        [TestMethod]
        public void ImplicitCast_ValidIndexer_SameReference()
        {
            var expected = new TableCell();
            var indexer = new TableCellIndexer(expected);
            var target = (TableCell)indexer;

            Assert.AreSame(expected, target);
        }

        [TestMethod]
        public void ImplicitCast_Null_IsNull()
        {
            TableCellIndexer indexer = null;
            var target = (TableCell)indexer;
            Assert.IsNull(target);
        }

        #endregion
    }
}
