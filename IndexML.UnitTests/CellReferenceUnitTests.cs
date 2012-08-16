namespace IndexML.UnitTests
{
    using System;
    using DocumentFormat.OpenXml.Spreadsheet;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Unit tests for the <see cref="CellReference"/> class.
    /// </summary>
    [TestClass]
    public class CellReferenceUnitTests : OpenXmlIndexerTest
    {
        #region Test Methods

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Create_NullCellArgument_ExceptionThrown()
        {
            var target = CellReference.Create((Cell)null);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void Create_NullStringArgument_ExceptionThrown()
        {
            var target = CellReference.Create((string)null);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void Create_MalformedCellReference_ExceptionThrown()
        {
            var target = CellReference.Create(new Cell() { CellReference = string.Empty });
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void Create_MalformedCellReferenceValue_ExceptionThrown()
        {
            var target = CellReference.Create(new Cell() { CellReference = null });
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void Create_EmptyString_ExceptionThrown()
        {
            var target = CellReference.Create(string.Empty);
        }

        [TestMethod]
        public void Create_ValidSingleCell_ObjectCreated()
        {
            var target = CellReference.Create(new Cell() { CellReference = "A1" });
            Assert.IsNotNull(target);
        }

        [TestMethod]
        public void IsValidCellReference_NullString_ReturnsFalse()
        {
            Assert.IsFalse(CellReference.IsValidCellReference(null));
        }

        [TestMethod]
        public void IsSingleCellReference_NullString_ReturnsFalse()
        {
            Assert.IsFalse(CellReference.IsSingleCellReference(null));
        }

        [TestMethod]
        public void IsRangeCellReference_NullString_ReturnsFalse()
        {
            Assert.IsFalse(CellReference.IsRangeCellReference(null));
        }

        [TestMethod]
        public void TryGetRowIndex_ValidString_ReturnsTrue()
        {
            long rowIdx;
            var result = CellReference.TryGetRowIndex("A1", out rowIdx);

            Assert.IsTrue(result);
            Assert.AreEqual(1, rowIdx);
        }

        [TestMethod]
        public void TryGetRowIndex_MalformedRowIndex_ReturnsFalse()
        {
            long rowIdx;
            var result = CellReference.TryGetRowIndex("A$1", out rowIdx);
            
            Assert.IsFalse(result);
            Assert.AreEqual(default(long), rowIdx);
        }

        [TestMethod]
        public void TryGetColumnName_StrictValidString_ReturnsTrue()
        {
            string colName;
            var result = CellReference.TryGetColumnName("A1", true, out colName);
            
            Assert.IsTrue(result);
            Assert.AreEqual("A", colName, true);
        }

        [TestMethod]
        public void TryGetColumnName_StrictInvalidString_ReturnsFalse()
        {
            string colName;
            var result = CellReference.TryGetColumnName("A", true, out colName);

            Assert.IsFalse(result);
            Assert.AreEqual(default(string), colName, true);
        }

        [TestMethod]
        public void TryGetColumnName_NotStrictValidString_ReturnsTrue()
        {
            string colName;
            var result = CellReference.TryGetColumnName("A", false, out colName);

            Assert.IsTrue(result);
            Assert.AreEqual("A", colName, true);
        }

        [TestMethod]
        public void TryGetColumnName_NotStrictInvalidString_ReturnsFalse()
        {
            string colName;
            var result = CellReference.TryGetColumnName("A$", false, out colName);

            Assert.IsFalse(result);
            Assert.AreEqual(default(string), colName, true);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentOutOfRangeException))]
        public void GetColumnName_NegativeColumn_ThrowsException()
        {
            var name = CellReference.GetColumnName(-1);
        }

        [TestMethod]
        public void GetColumnName_ValidColumnIndex_CorrectName()
        {
            var A = CellReference.GetColumnName(1);
            Assert.AreEqual("A", A, true);

            var AA = CellReference.GetColumnName(27);
            Assert.AreEqual("AA", AA, true);
        }

        [TestMethod]
        public void TryGetColumnIndex_StrictValidSingleCell_ReturnsExpected()
        {
            long colIdx;
            var result = CellReference.TryGetColumnIndex("A1", true, out colIdx);

            Assert.IsTrue(result);
            Assert.AreEqual(1, colIdx);
        }

        [TestMethod]
        public void TryGetColumnIndex_NotStrictValidSingleCell_ReturnsExpected()
        {
            long colIdx;
            var result = CellReference.TryGetColumnIndex("A", false, out colIdx);

            Assert.IsTrue(result);
            Assert.AreEqual(1, colIdx);
        }

        [TestMethod]
        public void TryGetColumnIndex_StrictDoubleLetterCell_ReturnsExpected()
        {
            long colIdx;
            var result = CellReference.TryGetColumnIndex("AZ1", true, out colIdx);

            Assert.IsTrue(result);
            Assert.AreEqual(52, colIdx);
        }

        [TestMethod]
        public void TryGetcolumnIndex_NotStrictDoubleLetterCell_ReturnsExpected()
        {
            long colIdx;
            var result = CellReference.TryGetColumnIndex("AZ", false, out colIdx);

            Assert.IsTrue(result);
            Assert.AreEqual(52, colIdx);
        }

        [TestMethod]
        public void TryGetColumnIndex_StrictTripleLetterCell_ReturnsExpected()
        {
            long colIdx;
            var result = CellReference.TryGetColumnIndex("AAA1", true, out colIdx);

            Assert.IsTrue(result);
            Assert.AreEqual(703, colIdx);
        }

        [TestMethod]
        public void TryGetcolumnIndex_NotStrictTripleLetterCell_ReturnsExpected()
        {
            long colIdx;
            var result = CellReference.TryGetColumnIndex("AAA", false, out colIdx);

            Assert.IsTrue(result);
            Assert.AreEqual(703, colIdx);
        }

        [TestMethod]
        public void ToString_ValidCell_ReturnsValue()
        {
            string refStr = "A1";
            var target = CellReference.Create(refStr);            
            Assert.AreEqual(refStr, target.ToString());
        }

        [TestMethod]
        public void ValueEquals_EqualCells_ReturnsTrue()
        {
            var target = CellReference.Create("A1");
            var other = CellReference.Create("A1");

            Assert.IsTrue(CellReference.ValueEquals(target, other));
        }

        [TestMethod]
        public void ValueEquals_NonEqualCells_ReturnsFalse()
        {
            var target = CellReference.Create("A1");
            var other = CellReference.Create("A2");

            Assert.IsFalse(CellReference.ValueEquals(target, other));
        }

        #endregion
    }
}
