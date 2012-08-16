namespace IndexML.UnitTests
{
    using System;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Unit tests for the <see cref="SingleCellReference"/> class.
    /// </summary>
    [TestClass]
    public class SingleCellReferenceUnitTests : OpenXmlIndexerTest
    {
        #region Test Methods

        #region Constructor

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void Constructor_NullArgument_ThrowsException()
        {
            var target = new SingleCellReference(null);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void Constructor_EmptyString_ThrowsException()
        {
            var target = new SingleCellReference(string.Empty);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void Constructor_MalformedCellReference_ThrowsException()
        {
            var target = new SingleCellReference("A$32");
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void Constructor_RangeCellRef_ThrowsException()
        {
            var target = new SingleCellReference("A2:B4");
        }

        [TestMethod]
        public void Constructor_ValidCell_CorrectProperties()
        {
            string col = "A";
            long idx = 1;
            var reference = col.ToString() + idx.ToString();

            var target = new SingleCellReference(reference);

            Assert.AreEqual(reference, target.Value, true);
            Assert.AreEqual("A", target.ColumnName, true);
            Assert.AreEqual(1, target.ColumnIndex);
            Assert.AreEqual(1, target.RowIndex);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentOutOfRangeException))]
        public void Constructor_ZeroRows_ThrowsException()
        {
            var target = new SingleCellReference(1, 0);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentOutOfRangeException))]
        public void Constructor_ZeroColumns_ThrowsException()
        {
            var target = new SingleCellReference(0, 1);
        }

        [TestMethod]
        public void Constructor_InRangeRowAndCol_NoError()
        {
            var target = new SingleCellReference(1, 1);

            Assert.AreEqual(1, target.ColumnIndex);
            Assert.AreEqual(1, target.RowIndex);
            Assert.AreEqual("A", target.ColumnName, true);
            Assert.AreEqual("A1", target.Value, true);
        }

        #endregion

        #region ContainsOrSubsumes
        
        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void ContainsOrSubsumes_NullArgument_ThrowsException()
        {
            var target = new SingleCellReference("A1");
            target.ContainsOrSubsumes(null);
        }

        [TestMethod]
        public void ContainsOrSubsumes_DifferentCell_ReturnsFalse()
        {
            var target = new SingleCellReference("A1");
            var other = new SingleCellReference("A2");

            Assert.IsFalse(target.ContainsOrSubsumes(other));
        }

        [TestMethod]
        public void ContainsOrSubsumes_SameCell_ReturnsTrue()
        {
            var target = new SingleCellReference("A1");
            var other = new SingleCellReference("A1");

            Assert.IsTrue(target.ContainsOrSubsumes(other));
        }

        #endregion

        #region Translate

        [TestMethod]
        public void Translate_Zero_NoChanges()
        {
            var target = new SingleCellReference("C1");
            var result = target.Translate(0, 0);

            Assert.IsNotNull(result);
            Assert.IsTrue(CellReference.ValueEquals(target, result));
        }

        [TestMethod]
        public void Translate_PositiveRows_ReturnsCorrectCell()
        {
            var target = new SingleCellReference("A1");
            var result = target.Translate(3, 0) as SingleCellReference;

            // Should still be a single cell            
            Assert.IsNotNull(result);

            // Rows should be translated
            Assert.AreEqual(4, result.RowIndex);

            // Columns should be unchanged
            Assert.AreEqual(target.ColumnIndex, result.ColumnIndex);
            Assert.AreEqual(target.ColumnName, result.ColumnName, true);
        }

        [TestMethod]
        public void Translate_PositiveColumns_ReturnsCorrectCell()
        {
            var target = new SingleCellReference("A1");
            var result = target.Translate(0, 3) as SingleCellReference;

            // Should still be a single cell            
            Assert.IsNotNull(result);

            // Rows should be unchanged
            Assert.AreEqual(target.RowIndex, result.RowIndex);

            // Columns should be translated
            Assert.AreEqual(4, result.ColumnIndex);
            Assert.AreEqual("D", result.ColumnName, true);
        }

        [TestMethod]
        public void Translate_NegativeRowsNotPastOrigin_ReturnsCorrectCell()
        {
            var target = new SingleCellReference("A4");
            var result = target.Translate(-3, 0) as SingleCellReference;

            // Should still be a single cell            
            Assert.IsNotNull(result);

            // Rows should be translated
            Assert.AreEqual(1, result.RowIndex);

            // Columns should be unchanged
            Assert.AreEqual(target.ColumnName, result.ColumnName, true);
            Assert.AreEqual(target.ColumnName, result.ColumnName, true);
        }

        [TestMethod]
        public void Translate_NegativeRowsPastOrigin_ReturnsOrigin()
        {
            var target = new SingleCellReference("A4");
            var result = target.Translate(-5, 0) as SingleCellReference;

            // Should still be a single cell            
            Assert.IsNotNull(result);

            // Rows should be translated
            Assert.AreEqual(1, result.RowIndex);

            // Columns should be unchanged
            Assert.AreEqual(target.ColumnName, result.ColumnName, true);
            Assert.AreEqual(target.ColumnName, result.ColumnName, true);
        }

        [TestMethod]
        public void Translate_NegativeColumnsNotPastOrigin_ReturnsCorrectCell()
        {
            var target = new SingleCellReference("D1");
            var result = target.Translate(0, -3) as SingleCellReference;

            // Should still be a single cell            
            Assert.IsNotNull(result);

            // Rows should be unchanged
            Assert.AreEqual(target.RowIndex, result.RowIndex);

            // Columns should be translated
            Assert.AreEqual(1, result.ColumnIndex);
            Assert.AreEqual("A", result.ColumnName, true);
        }

        [TestMethod]
        public void Translate_NegativeColumnsPastOrigin_ReturnsOrigin()
        {
            var target = new SingleCellReference("D1");
            var result = target.Translate(0, -5) as SingleCellReference;

            // Should still be a single cell            
            Assert.IsNotNull(result);

            // Rows should be unchanged
            Assert.AreEqual(target.RowIndex, result.RowIndex);

            // Columns should be translated
            Assert.AreEqual(1, result.ColumnIndex);
            Assert.AreEqual("A", result.ColumnName, true);
        }

        #endregion

        #region Scale

        [TestMethod]
        public void Scale_Zero_NoChanges()
        {
            var target = new SingleCellReference("D4");
            var result = target.Scale(0, 0);

            Assert.IsNotNull(result);
            Assert.IsTrue(CellReference.ValueEquals(target, result));
        }

        [TestMethod]
        public void Scale_PositiveRows_ReturnsCorrectRange()
        {
            var target = new SingleCellReference("A1");
            var result = target.Scale(3, 0) as RangeCellReference;

            // Should be a range
            Assert.IsNotNull(result);

            // Starting cell row should be equal to target
            Assert.IsTrue(CellReference.ValueEquals(target, result.StartingCellReference));
            Assert.AreEqual(target.ColumnName, result.StartingCellReference.ColumnName, true);
            Assert.AreEqual(target.ColumnIndex, result.StartingCellReference.ColumnIndex);
            Assert.AreEqual(target.RowIndex, result.StartingCellReference.RowIndex);

            // Ending cell should be translated with column unchanged
            Assert.IsFalse(CellReference.ValueEquals(target, result.EndingCellReference));
            Assert.AreEqual(target.ColumnName, result.EndingCellReference.ColumnName, true);
            Assert.AreEqual(target.ColumnIndex, result.EndingCellReference.ColumnIndex);
            Assert.AreEqual(4, result.EndingCellReference.RowIndex);
        }

        [TestMethod]
        public void Scale_PositiveColumns_ReturnsCorrectRange()
        {
            var target = new SingleCellReference("A1");
            var result = target.Scale(0, 3) as RangeCellReference;

            // Should be a range
            Assert.IsNotNull(result);

            // Starting cell row should be equal to target
            Assert.IsTrue(CellReference.ValueEquals(target, result.StartingCellReference));
            Assert.AreEqual(target.ColumnName, result.StartingCellReference.ColumnName, true);
            Assert.AreEqual(target.ColumnIndex, result.StartingCellReference.ColumnIndex);
            Assert.AreEqual(target.RowIndex, result.StartingCellReference.RowIndex);

            // Ending cell should be translated with row unchanged
            Assert.IsFalse(CellReference.ValueEquals(target, result.EndingCellReference));
            Assert.AreEqual("D", result.EndingCellReference.ColumnName, true);
            Assert.AreEqual(4, result.EndingCellReference.ColumnIndex);
            Assert.AreEqual(target.RowIndex, result.EndingCellReference.RowIndex);
        }

        [TestMethod]
        public void Scale_NegativeRows_ReturnsCorrectRange()
        {
            var target = new SingleCellReference("A4");
            var result = target.Scale(-3, 0) as RangeCellReference;

            // Should be a range
            Assert.IsNotNull(result);

            // Starting cell row should be translated with column unchanged
            Assert.IsFalse(CellReference.ValueEquals(target, result.StartingCellReference)); 
            Assert.AreEqual(target.ColumnName, result.StartingCellReference.ColumnName, true);
            Assert.AreEqual(target.ColumnIndex, result.StartingCellReference.ColumnIndex);
            Assert.AreEqual(1, result.StartingCellReference.RowIndex);

            // Ending cell should be equal to target
            Assert.IsTrue(CellReference.ValueEquals(target, result.EndingCellReference)); 
            Assert.AreEqual(target.ColumnName, result.EndingCellReference.ColumnName, true);
            Assert.AreEqual(target.ColumnIndex, result.EndingCellReference.ColumnIndex);
            Assert.AreEqual(target.RowIndex, result.EndingCellReference.RowIndex);
        }

        [TestMethod]
        public void Scale_NegativeRowsPastOrigin_ReturnsCorrectRange()
        {
            var target = new SingleCellReference("A4");
            var result = target.Scale(-5, 0) as RangeCellReference;

            // Should be a range
            Assert.IsNotNull(result);

            // Starting cell row should be translated with column unchanged
            Assert.IsFalse(CellReference.ValueEquals(target, result.StartingCellReference));
            Assert.AreEqual(target.ColumnName, result.StartingCellReference.ColumnName, true);
            Assert.AreEqual(target.ColumnIndex, result.StartingCellReference.ColumnIndex);
            Assert.AreEqual(1, result.StartingCellReference.RowIndex);

            // Ending cell should be equal to target
            Assert.IsTrue(CellReference.ValueEquals(target, result.EndingCellReference));
            Assert.AreEqual(target.ColumnName, result.EndingCellReference.ColumnName, true);
            Assert.AreEqual(target.ColumnIndex, result.EndingCellReference.ColumnIndex);
            Assert.AreEqual(target.RowIndex, result.EndingCellReference.RowIndex);
        }

        [TestMethod]
        public void Scale_NegativeRowsAtOrigin_ReturnsOrigin()
        {
            var target = new SingleCellReference("A1");
            var result = target.Scale(-1, 0) as SingleCellReference;

            // Should be a single cell
            Assert.IsNotNull(result);

            // Values should be equal
            Assert.IsTrue(CellReference.ValueEquals(target, result));

            // Row indexes should be equal
            Assert.AreEqual(target.RowIndex, result.RowIndex);

            // Columns should be equal
            Assert.AreEqual(target.ColumnName, result.ColumnName, true);
            Assert.AreEqual(target.ColumnIndex, result.ColumnIndex);
        }

        [TestMethod]
        public void Scale_NegativeColumns_ReturnsCorrectRange()
        {
            var target = new SingleCellReference("D1");
            var result = target.Scale(0, -3) as RangeCellReference;

            // Should be a range
            Assert.IsNotNull(result);

            // Starting cell row should be translated with row unchanged
            Assert.IsFalse(CellReference.ValueEquals(target, result.StartingCellReference));
            Assert.AreEqual("A", result.StartingCellReference.ColumnName, true);
            Assert.AreEqual(1, result.StartingCellReference.ColumnIndex);
            Assert.AreEqual(target.RowIndex, result.StartingCellReference.RowIndex);

            // Ending cell should be equal to target
            Assert.IsTrue(CellReference.ValueEquals(target, result.EndingCellReference));
            Assert.AreEqual(target.ColumnName, result.EndingCellReference.ColumnName, true);
            Assert.AreEqual(target.ColumnIndex, result.EndingCellReference.ColumnIndex);
            Assert.AreEqual(target.RowIndex, result.EndingCellReference.RowIndex);
        }

        [TestMethod]
        public void Scale_NegativeColumnsPastOrigin_ReturnsCorrectRange()
        {
            var target = new SingleCellReference("D1");
            var result = target.Scale(0, -5) as RangeCellReference;

            // Should be a range
            Assert.IsNotNull(result);

            // Starting cell row should be translated with row unchanged
            Assert.IsFalse(CellReference.ValueEquals(target, result.StartingCellReference));
            Assert.AreEqual("A", result.StartingCellReference.ColumnName, true);
            Assert.AreEqual(1, result.StartingCellReference.ColumnIndex);
            Assert.AreEqual(target.RowIndex, result.StartingCellReference.RowIndex);

            // Ending cell should be equal to target
            Assert.IsTrue(CellReference.ValueEquals(target, result.EndingCellReference));
            Assert.AreEqual(target.ColumnName, result.EndingCellReference.ColumnName, true);
            Assert.AreEqual(target.ColumnIndex, result.EndingCellReference.ColumnIndex);
            Assert.AreEqual(target.RowIndex, result.EndingCellReference.RowIndex);
        }

        [TestMethod]
        public void Scale_NegativeColumnsAtOrigin_ReturnsOrigin()
        {
            var target = new SingleCellReference("A1");
            var result = target.Scale(0, -1) as SingleCellReference;

            // Should be a single cell
            Assert.IsNotNull(result);

            // Values should be equal
            Assert.IsTrue(CellReference.ValueEquals(target, result));

            // Row indexes should be equal
            Assert.AreEqual(target.RowIndex, result.RowIndex);

            // Columns should be equal
            Assert.AreEqual(target.ColumnName, result.ColumnName, true);
            Assert.AreEqual(target.ColumnIndex, result.ColumnIndex);
        }

        #endregion

        #endregion
    }
}
