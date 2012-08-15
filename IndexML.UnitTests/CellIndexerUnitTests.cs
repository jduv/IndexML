namespace IndexML.UnitTests
{
    using System;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using DocumentFormat.OpenXml.Spreadsheet;
    using DocumentFormat.OpenXml;

    [TestClass]
    public class CellIndexerUnitTests
    {
        #region Test Methods

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Constructor_NullArgument_ThrowsException()
        {
            var target = new CellIndexer(null);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void ColumnIndex_NullCellReference_ThrowsException()
        {
            var cell = new Cell()
            {
                CellReference = null
            };

            var target = new CellIndexer(cell);
            var explode = target.ColumnIndex;
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void ColumnName_NullCellReference_ThrowsException()
        {
            var cell = new Cell()
            {
                CellReference = null
            };

            var target = new CellIndexer(cell);
            var explode = target.ColumnName;
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void ColumnIndex_MalformedReference_ThrowsException()
        {
            var cell = new Cell()
            {
                CellReference = new StringValue() {  Value = "A#@%!1" }
            };

            var target = new CellIndexer(cell);
            var explode = target.ColumnIndex;
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void ColumnName_MalformedReference_ThrowsException()
        {
            var cell = new Cell()
            {
                CellReference = new StringValue() { Value = "A#@%!1" }
            };

            var target = new CellIndexer(cell);
            var explode = target.ColumnName;
        }

        [TestMethod]
        public void ColumnIndex_ValidCellReference_CorrectColumnIndex()
        {
            var cell = new Cell()
            {
                CellReference = new StringValue() { Value = "A1" }
            };

            var target = new CellIndexer(cell);
            Assert.AreEqual(1, target.ColumnIndex);
        }

        [TestMethod]
        public void ColumnName_ValidCellReference_CorrectColumnName()
        {
            var cell = new Cell()
            {
                CellReference = new StringValue() { Value = "A1" }
            };

            var target = new CellIndexer(cell);
            Assert.AreEqual("A", target.ColumnName);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void GetColumnIndex_NullCell_ThrowsException()
        {
            var explode = CellIndexer.GetColumnIndex((Cell)null);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void GetColumnIndex_NullCellReference_ThrowsException()
        {
            var explode = CellIndexer.GetColumnIndex(new Cell() { CellReference = null });
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void GetColumnIndex_NullStringValueInCellReference_ThrowsException()
        {
            var explode = CellIndexer.GetColumnIndex(new Cell()
            {
                CellReference = new StringValue() { Value = null }
            });
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void GetColumnIndex_EmptyStringInCellReference_ThrowsException()
        {
            var explode = CellIndexer.GetColumnIndex(new Cell()
            {
                CellReference = new StringValue() { Value = string.Empty }
            });
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void GetColumnIndex_MalformedCellReference_ThrowsException()
        {
            var explode = CellIndexer.GetColumnIndex(new Cell() 
            {
                CellReference = new StringValue() { Value = "A@#$%g1" }
            });
        }

        [TestMethod]        
        public void GetColumnIndex_ValidCellReference_CorrectColumnIndex()
        {
            var result = CellIndexer.GetColumnIndex(new Cell()
            {
                CellReference = new StringValue() { Value = "Z123" }
            });

            Assert.AreEqual(26, result);
        }

        #endregion
    }
}
