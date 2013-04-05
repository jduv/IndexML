namespace IndexML.UnitTests.Wordprocessing
{
    using System;
    using System.Linq;
    using DocumentFormat.OpenXml.Wordprocessing;
    using IndexML.Wordprocessing;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    
    [TestClass]
    [DeploymentItem(@"IndexML.TestFiles\", @"IndexML.TestFiles\")]
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
        [ExpectedException(typeof(ArgumentException))]
        public void Constructor_NoCells_ThrowsException()
        {
            var target = new TableRowIndexer(new TableRow());
        }

        [TestMethod]
        public void Constructor_ValidDoc_ValidState()
        {
            SafeExecuteTest(
                StandardDocPath,
                (doc) =>
                {
                    var table = doc.MainDocumentPart.Document.Body.Elements<Table>().FirstOrDefault();
                    var expected = table.Elements<TableRow>().FirstOrDefault();
                    var target = new TableRowIndexer(expected);

                    Assert.IsNotNull(target.Row);
                    Assert.IsNotNull(target.Cells);
                    Assert.AreEqual(2, target.Cells.Count()); // Two cols, check doc to verify.
                });
        }

        [TestMethod]
        public void ImplicitCast_ValidIndexer_SameReference()
        {
            var expected = new TableRow();
            expected.AppendChild<TableCell>(new TableCell());

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
