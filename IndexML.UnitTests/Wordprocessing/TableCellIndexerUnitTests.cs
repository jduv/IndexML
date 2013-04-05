namespace IndexML.UnitTests.Wordprocessing
{
    using System;
    using System.Linq;
    using DocumentFormat.OpenXml.Wordprocessing;
    using IndexML.Wordprocessing;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [DeploymentItem(@"IndexML.TestFiles\", @"IndexML.TestFiles\")]
    public class TableCellIndexerUnitTests : WordprocessingDocumentTest
    {
        #region Test Methods

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Constructor_NullArgument_ThrowsException()
        {
            var target = new TableCellIndexer(null);
        }

        [TestMethod]
        public void Constructor_ValidTableCell_ValidState()
        {
            SafeExecuteTest(
                ComplexSingleTableCell,
                (doc) =>
                {
                    var expectedCell = doc.MainDocumentPart.Document.Body.Elements<Table>().FirstOrDefault()
                        .Descendants<TableCell>().FirstOrDefault();
                    var indexer = new TableCellIndexer(expectedCell);

                    Assert.IsNotNull(indexer.Cell);
                    Assert.AreSame(expectedCell, indexer.Cell);
                    Assert.IsNotNull(indexer.Paragraphs);
                    Assert.IsNotNull(indexer.Tables);
                });
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
