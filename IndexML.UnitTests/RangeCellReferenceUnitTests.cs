namespace IndexML.UnitTests
{
    using System;
    using IndexML.Spreadsheet;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Moq;

    /// <summary>
    /// Unit tests for the <see cref="RangeCellReference"/> class.
    /// </summary>
    [TestClass]
    public class RangeCellReferenceUnitTests : OpenXmlIndexerTest
    {
        #region Test Methods

        #region Constructor

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void Constructor_NullArgument_ThrowsException()
        {
            var target = new RangeCellReference(null);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void Constructor_EmptyString_ThrowsException()
        {
            var target = new RangeCellReference(string.Empty);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void Constructor_MalformedCellReference_ThrowsException()
        {
            var target = new RangeCellReference("A$32");
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void Constructor_SingleCellRef_ThrowsException()
        {
            var target = new RangeCellReference("A1");
        }

        [TestMethod]
        public void Constructor_ValidCellRange_NoError()
        {
            var target = new RangeCellReference("A1:C4");

            Assert.AreEqual("A", target.StartingCellReference.ColumnName, true);
            Assert.AreEqual(1, target.StartingCellReference.ColumnIndex);
            Assert.AreEqual(1, target.StartingCellReference.RowIndex);

            Assert.AreEqual("C", target.EndingCellReference.ColumnName, true);
            Assert.AreEqual(3, target.EndingCellReference.ColumnIndex);
            Assert.AreEqual(4, target.EndingCellReference.RowIndex);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Constructor_NullStartingCell_ThrowsException()
        {
            var target = new RangeCellReference(null, new SingleCellReference("A2"));
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Constructor_NullEndingCell_ThrowsException()
        {
            var target = new RangeCellReference(new SingleCellReference("A1"), null);
        }

        #endregion

        #region Properties

        [TestMethod]
        public void Rows_ValidCellRange_CorrectNumber()
        {
            var target = new RangeCellReference("A1:A3");
            Assert.AreEqual(3, target.Rows);
        }

        [TestMethod]
        public void Columns_ValidCellRange_CorrectNumber()
        {
            var target = new RangeCellReference("A1:C1");
            Assert.AreEqual(3, target.Columns);
        }

        #endregion

        #region ContainsOrSubsumes

        [TestMethod]
        public void ContainsOrSubsumes_SingleCellRef_NotContained()
        {
            var target = new RangeCellReference("A1:C4");
            var other = new SingleCellReference("E5");

            Assert.IsFalse(target.ContainsOrSubsumes(other));
        }

        [TestMethod]
        public void ContainsOrSubsumes_SingleCellRef_Contained()
        {
            var target = new RangeCellReference("A1:C4");
            var other = new SingleCellReference("B3");

            Assert.IsTrue(target.ContainsOrSubsumes(other));
        }

        [TestMethod]
        public void ContainsOrSubsumes_Range_NotContained()
        {
            var target = new RangeCellReference("A1:C4");
            var other = new RangeCellReference("D1:E5");

            Assert.IsFalse(target.ContainsOrSubsumes(other));
        }

        [TestMethod]
        public void ContainsOrSubsumes_Range_Contained()
        {
            var target = new RangeCellReference("A1:C4");
            var other = new RangeCellReference("B2:C4");

            Assert.IsTrue(target.ContainsOrSubsumes(other));
        }

        [TestMethod]
        public void ContainsOrSubsumes_OverlappingRange_NotContained()
        {
            var target = new RangeCellReference("A1:C4");
            var other = new RangeCellReference("B2:F5");

            Assert.IsFalse(target.ContainsOrSubsumes(other));
        }

        [TestMethod]        
        public void ContainsOrSubsumes_MoqCellReference_ThrowsException()
        {
            var target = new RangeCellReference("A1:C4");
            var mock = new Mock<ICellReference>();

            // Must be a known type
            Assert.IsFalse(target.ContainsOrSubsumes(mock.Object));
        }

        #endregion

        #region Translate

        [TestMethod]
        public void Translate_Zero_ReturnsOriginalRange()
        {
            var target = new RangeCellReference("D4:F16");
            var result = target.Move(0, 0);

            Assert.IsNotNull(result);
            Assert.IsTrue(CellReference.ValueEquals(target, result));            
        }

        [TestMethod]
        public void Translate_PositiveRows_ReturnsCorrectRange()
        {
            var target = new RangeCellReference("A1:C4");
            var result = target.Move(3, 0) as RangeCellReference;

            // Should still be a range            
            Assert.IsNotNull(result);

            // Rows should be translated
            Assert.AreEqual(4, result.StartingCellReference.RowIndex);
            Assert.AreEqual(7, result.EndingCellReference.RowIndex);

            // Columns should be unchanged
            Assert.AreEqual(target.StartingCellReference.ColumnName, result.StartingCellReference.ColumnName, true);
            Assert.AreEqual(target.EndingCellReference.ColumnName, result.EndingCellReference.ColumnName, true);
            Assert.AreEqual(target.StartingCellReference.ColumnIndex, result.StartingCellReference.ColumnIndex);
            Assert.AreEqual(target.EndingCellReference.ColumnIndex, result.EndingCellReference.ColumnIndex);
        }

        [TestMethod]
        public void Translate_PositiveColumns_ReturnsCorrectRange()
        {
            var target = new RangeCellReference("A1:C4");
            var result = target.Move(0, 3) as RangeCellReference;

            // Should still be a range            
            Assert.IsNotNull(result);

            // Rows should be unchanged
            Assert.AreEqual(target.StartingCellReference.RowIndex, result.StartingCellReference.RowIndex);
            Assert.AreEqual(target.EndingCellReference.RowIndex, result.EndingCellReference.RowIndex);

            // Columns should be translated
            Assert.AreEqual("D", result.StartingCellReference.ColumnName, true);
            Assert.AreEqual("F", result.EndingCellReference.ColumnName, true);
            Assert.AreEqual(4, result.StartingCellReference.ColumnIndex);
            Assert.AreEqual(6, result.EndingCellReference.ColumnIndex);
        }

        [TestMethod]
        public void Translate_NegativeRows_ReturnsCorrectRange()
        {
            var target = new RangeCellReference("A4:C7");
            var result = target.Move(-3, 0) as RangeCellReference;

            // Should still be a range            
            Assert.IsNotNull(result);

            // Rows should be translated
            Assert.AreEqual(1, result.StartingCellReference.RowIndex);
            Assert.AreEqual(4, result.EndingCellReference.RowIndex);

            // Columns should be unchanged
            Assert.AreEqual(target.StartingCellReference.ColumnName, result.StartingCellReference.ColumnName, true);
            Assert.AreEqual(target.EndingCellReference.ColumnName, result.EndingCellReference.ColumnName, true);
            Assert.AreEqual(target.StartingCellReference.ColumnIndex, result.StartingCellReference.ColumnIndex);
            Assert.AreEqual(target.EndingCellReference.ColumnIndex, result.EndingCellReference.ColumnIndex);
        }

        [TestMethod]
        public void Translate_NegativeRowsPastOrigin_ReturnsCorrectRange()
        {
            var target = new RangeCellReference("A4:C7");
            var result = target.Move(-5, 0) as RangeCellReference;

            // Should still be a range            
            Assert.IsNotNull(result);

            // Rows should be translated
            Assert.AreEqual(1, result.StartingCellReference.RowIndex);
            Assert.AreEqual(4, result.EndingCellReference.RowIndex);

            // Columns should be unchanged
            Assert.AreEqual(target.StartingCellReference.ColumnName, result.StartingCellReference.ColumnName, true);
            Assert.AreEqual(target.EndingCellReference.ColumnName, result.EndingCellReference.ColumnName, true);
            Assert.AreEqual(target.StartingCellReference.ColumnIndex, result.StartingCellReference.ColumnIndex);
            Assert.AreEqual(target.EndingCellReference.ColumnIndex, result.EndingCellReference.ColumnIndex);
        }

        [TestMethod]
        public void Translate_NegativeColumns_ReturnsCorrectRange()
        {
            var target = new RangeCellReference("D1:F4");
            var result = target.Move(0, -3) as RangeCellReference;

            // Should still be a range            
            Assert.IsNotNull(result);

            // Rows should be unchanged
            Assert.AreEqual(target.StartingCellReference.RowIndex, result.StartingCellReference.RowIndex);
            Assert.AreEqual(target.EndingCellReference.RowIndex, result.EndingCellReference.RowIndex);

            // Columns should be translated
            Assert.AreEqual("A", result.StartingCellReference.ColumnName, true);
            Assert.AreEqual("C", result.EndingCellReference.ColumnName, true);
            Assert.AreEqual(1, result.StartingCellReference.ColumnIndex);
            Assert.AreEqual(3, result.EndingCellReference.ColumnIndex);
        }

        [TestMethod]
        public void Translate_NegativeColumnsPastOrigin_ReturnsCorrectRange()
        {
            var target = new RangeCellReference("D1:F4");
            var result = target.Move(0, -5) as RangeCellReference;

            // Should still be a range            
            Assert.IsNotNull(result);

            // Rows should be unchanged
            Assert.AreEqual(target.StartingCellReference.RowIndex, result.StartingCellReference.RowIndex);
            Assert.AreEqual(target.EndingCellReference.RowIndex, result.EndingCellReference.RowIndex);

            // Columns should be translated
            Assert.AreEqual("A", result.StartingCellReference.ColumnName, true);
            Assert.AreEqual("C", result.EndingCellReference.ColumnName, true);
            Assert.AreEqual(1, result.StartingCellReference.ColumnIndex);
            Assert.AreEqual(3, result.EndingCellReference.ColumnIndex);
        }
        
        #endregion

        #region Scale

        [TestMethod]
        public void Scale_Zero_ReturnsOriginalRange()
        {
            var target = new RangeCellReference("D4:F16");
            var result = target.Resize(0, 0);

            Assert.IsNotNull(result);
            Assert.IsTrue(CellReference.ValueEquals(target, result));
        }

        [TestMethod]
        public void Scale_PositiveRows_ReturnsCorrectRange()
        {
            var target = new RangeCellReference("A1:C4");
            var result = target.Resize(3, 0) as RangeCellReference;

            // Should still be a range            
            Assert.IsNotNull(result);

            // Starting cell should be unchanged
            Assert.IsTrue(CellReference.ValueEquals(target.StartingCellReference, result.StartingCellReference));
            Assert.AreEqual(target.StartingCellReference.ColumnName, result.StartingCellReference.ColumnName, true);
            Assert.AreEqual(target.StartingCellReference.ColumnIndex, result.StartingCellReference.ColumnIndex);
            Assert.AreEqual(target.StartingCellReference.RowIndex, result.StartingCellReference.RowIndex);

            // Ending cell row should be scaled
            Assert.IsFalse(CellReference.ValueEquals(target.EndingCellReference, result.EndingCellReference));
            Assert.AreEqual(target.EndingCellReference.ColumnName, result.EndingCellReference.ColumnName, true);
            Assert.AreEqual(target.EndingCellReference.ColumnIndex, result.EndingCellReference.ColumnIndex);
            Assert.AreEqual(7, result.EndingCellReference.RowIndex);
        }

        [TestMethod]
        public void Scale_PositiveColumns_ReturnsCorrectRange()
        {
            var target = new RangeCellReference("A1:D4");
            var result = target.Resize(0, 3) as RangeCellReference;

            // Should still be a range            
            Assert.IsNotNull(result);

            // Starting cell should be unchanged
            Assert.IsTrue(CellReference.ValueEquals(target.StartingCellReference, result.StartingCellReference));
            Assert.AreEqual(target.StartingCellReference.ColumnName, result.StartingCellReference.ColumnName, true);
            Assert.AreEqual(target.StartingCellReference.ColumnIndex, result.StartingCellReference.ColumnIndex);
            Assert.AreEqual(target.StartingCellReference.RowIndex, result.StartingCellReference.RowIndex);

            // Ending cell column should be scaled
            Assert.IsFalse(CellReference.ValueEquals(target.EndingCellReference, result.EndingCellReference));
            Assert.AreEqual("G", result.EndingCellReference.ColumnName, true);
            Assert.AreEqual(7, result.EndingCellReference.ColumnIndex);
            Assert.AreEqual(target.EndingCellReference.RowIndex, result.EndingCellReference.RowIndex);
        }

        [TestMethod]
        public void Scale_NegativeRows_ReturnsCorrectRange()
        {
            var target = new RangeCellReference("A1:D4");
            var result = target.Resize(-3, 0) as RangeCellReference;

            // Should still be a range            
            Assert.IsNotNull(result);

            // Starting cell should be unchanged
            Assert.IsTrue(CellReference.ValueEquals(target.StartingCellReference, result.StartingCellReference));
            Assert.AreEqual(target.StartingCellReference.ColumnName, result.StartingCellReference.ColumnName, true);
            Assert.AreEqual(target.StartingCellReference.ColumnIndex, result.StartingCellReference.ColumnIndex);
            Assert.AreEqual(target.StartingCellReference.RowIndex, result.StartingCellReference.RowIndex);

            // Ending cell row should be scaled
            Assert.IsFalse(CellReference.ValueEquals(target.EndingCellReference, result.EndingCellReference));
            Assert.AreEqual(target.EndingCellReference.ColumnName, result.EndingCellReference.ColumnName, true);
            Assert.AreEqual(target.EndingCellReference.ColumnIndex, result.EndingCellReference.ColumnIndex);
            Assert.AreEqual(1, result.EndingCellReference.RowIndex);
        }

        [TestMethod]
        public void Scale_NegativeRowsPastOrigin_ReturnsCorrectRange()
        {
            var target = new RangeCellReference("A1:D4");
            var result = target.Resize(-5, 0) as RangeCellReference;

            // Should still be a range            
            Assert.IsNotNull(result);

            // Starting cell should be unchanged
            Assert.IsTrue(CellReference.ValueEquals(target.StartingCellReference, result.StartingCellReference));
            Assert.AreEqual(target.StartingCellReference.ColumnName, result.StartingCellReference.ColumnName, true);
            Assert.AreEqual(target.StartingCellReference.ColumnIndex, result.StartingCellReference.ColumnIndex);
            Assert.AreEqual(target.StartingCellReference.RowIndex, result.StartingCellReference.RowIndex);

            // Ending cell row should be scaled
            Assert.IsFalse(CellReference.ValueEquals(target.EndingCellReference, result.EndingCellReference));
            Assert.AreEqual(target.EndingCellReference.ColumnName, result.EndingCellReference.ColumnName, true);
            Assert.AreEqual(target.EndingCellReference.ColumnIndex, result.EndingCellReference.ColumnIndex);
            Assert.AreEqual(1, result.EndingCellReference.RowIndex);
        }

        [TestMethod]
        public void Scale_NegativeRowsPastStartingCell_ReturnsCorrectRange()
        {
            var target = new RangeCellReference("D4:F6");
            var result = target.Resize(-4, 0) as RangeCellReference;

            // Should be a range
            Assert.IsNotNull(result);

            // Starting cell should be changed
            Assert.IsFalse(CellReference.ValueEquals(target.StartingCellReference, result.StartingCellReference));
            Assert.AreEqual("D", result.StartingCellReference.ColumnName, true);
            Assert.AreEqual(4, result.StartingCellReference.ColumnIndex);
            Assert.AreEqual(2, result.StartingCellReference.RowIndex);

            // Ending cell should be changed
            Assert.IsFalse(CellReference.ValueEquals(target.StartingCellReference, result.StartingCellReference));
            Assert.AreEqual("F", result.EndingCellReference.ColumnName, true);
            Assert.AreEqual(6, result.EndingCellReference.ColumnIndex);
            Assert.AreEqual(4, result.EndingCellReference.RowIndex);
        }

        [TestMethod]
        public void Scale_NegativeColumns_ReturnsCorrectRange()
        {
            var target = new RangeCellReference("A1:D4");
            var result = target.Resize(0, -3) as RangeCellReference;

            // Should still be a range            
            Assert.IsNotNull(result);

            // Starting cell should be unchanged
            Assert.IsTrue(CellReference.ValueEquals(target.StartingCellReference, result.StartingCellReference));
            Assert.AreEqual(target.StartingCellReference.ColumnName, result.StartingCellReference.ColumnName, true);
            Assert.AreEqual(target.StartingCellReference.ColumnIndex, result.StartingCellReference.ColumnIndex);
            Assert.AreEqual(target.StartingCellReference.RowIndex, result.StartingCellReference.RowIndex);

            // Ending cell column should be scaled
            Assert.IsFalse(CellReference.ValueEquals(target.EndingCellReference, result.EndingCellReference));
            Assert.AreEqual("A", result.EndingCellReference.ColumnName, true);
            Assert.AreEqual(1, result.EndingCellReference.ColumnIndex);
            Assert.AreEqual(target.EndingCellReference.RowIndex, result.EndingCellReference.RowIndex);
        }       

        [TestMethod]
        public void Scale_NegativeColumnsPastOrigin_ReturnsCorrectRange()
        {
            var target = new RangeCellReference("A1:D4");
            var result = target.Resize(0, -5) as RangeCellReference;

            // Should still be a range            
            Assert.IsNotNull(result);

            // Starting cell should be unchanged
            Assert.IsTrue(CellReference.ValueEquals(target.StartingCellReference, result.StartingCellReference));
            Assert.AreEqual(target.StartingCellReference.ColumnName, result.StartingCellReference.ColumnName, true);
            Assert.AreEqual(target.StartingCellReference.ColumnIndex, result.StartingCellReference.ColumnIndex);
            Assert.AreEqual(target.StartingCellReference.RowIndex, result.StartingCellReference.RowIndex);

            // Ending cell column should be scaled
            Assert.IsFalse(CellReference.ValueEquals(target.EndingCellReference, result.EndingCellReference));
            Assert.AreEqual("A", result.EndingCellReference.ColumnName, true);
            Assert.AreEqual(1, result.EndingCellReference.ColumnIndex);
            Assert.AreEqual(target.EndingCellReference.RowIndex, result.EndingCellReference.RowIndex);
        }

        [TestMethod]
        public void Scale_NegativeColumnsPastStartingCell_ReturnsCorrectRange()
        {
            var target = new RangeCellReference("D4:F6");
            var result = target.Resize(0, -4) as RangeCellReference;

            // Should be a range
            Assert.IsNotNull(result);

            // Starting cell should be changed
            Assert.IsFalse(CellReference.ValueEquals(target.StartingCellReference, result.StartingCellReference));
            Assert.AreEqual("B", result.StartingCellReference.ColumnName, true);
            Assert.AreEqual(2, result.StartingCellReference.ColumnIndex);
            Assert.AreEqual(4, result.StartingCellReference.RowIndex);

            // Ending cell should be changed
            Assert.IsFalse(CellReference.ValueEquals(target.StartingCellReference, result.StartingCellReference));
            Assert.AreEqual("D", result.EndingCellReference.ColumnName, true);
            Assert.AreEqual(4, result.EndingCellReference.ColumnIndex);
            Assert.AreEqual(6, result.EndingCellReference.RowIndex);
        }

        [TestMethod]
        public void Scale_CollapseRange_ReturnsSingleCell()
        {
            var target = new RangeCellReference("B2:D4");
            var result = target.Resize(-2, -2) as SingleCellReference;

            // Should be a single cell
            Assert.IsNotNull(result);

            // Should be cell B2
            Assert.AreEqual("B", result.ColumnName);
            Assert.AreEqual(2, result.ColumnIndex);
            Assert.AreEqual(2, result.RowIndex);
        }

        [TestMethod]
        public void Scale_CollapseToOrigin_ReturnsOrigin()
        {
            var target = new RangeCellReference("A1:D4");
            var result = target.Resize(-4, -4) as SingleCellReference;

            // Should be a single cell
            Assert.IsNotNull(result);

            // Should be cell B2
            Assert.AreEqual("A", result.ColumnName);
            Assert.AreEqual(1, result.ColumnIndex);
            Assert.AreEqual(1, result.RowIndex);
        }

        #endregion

        #endregion
    }
}
