namespace IndexML.UnitTests
{
    using System;
    using DocumentFormat.OpenXml;
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
        public void ToString_ValidCell_ReturnsValue()
        {
            string refStr = "A1";
            var target = CellReference.Create(refStr);
            Assert.AreEqual(refStr, target.ToString());
        }

        #endregion
    }
}
