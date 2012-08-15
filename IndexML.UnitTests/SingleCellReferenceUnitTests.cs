namespace IndexML.UnitTests
{
    using System;
    using System.Text;
    using System.Collections.Generic;
    using System.Linq;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Unit tests for the <see cref="SingleCellReference"/> class.
    /// </summary>
    [TestClass]
    public class SingleCellReferenceUnitTests : OpenXmlIndexerTest
    {
        #region Test Methods

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
        [ExpectedException(typeof(ArgumentNullException))]
        public void ContainsOrSubsumes_NullArgument_ThrowsException()
        {
            var target = new SingleCellReference("A1");
            target.ContainsOrSubsumes(null);
        }

        [TestMethod]
        public void ContainsOrSubsumes_DifferentCell_False()
        {
            var target = new SingleCellReference("A1");
            var other = new SingleCellReference("A2");

            Assert.IsFalse(target.ContainsOrSubsumes(other));
        }

        [TestMethod]
        public void ContainsOrSubsumes_SameCell_True()
        {
            var target = new SingleCellReference("A1");
            var other = new SingleCellReference("A1");

            Assert.IsTrue(target.ContainsOrSubsumes(other));
        }

        [TestMethod]
        public void ExtendColumnRange_PositiveArgument_CorrectRange()
        {
            var target = new SingleCellReference("A1");
            var result = target.ExtendColumnRange(3) as RangeCellReference;

            Assert.AreEqual(target.ColumnName, result.StartingCellReference.ColumnName, true);
            Assert.AreEqual(target.ColumnIndex, result.StartingCellReference.ColumnIndex);
            Assert.AreEqual(target.RowIndex, result.StartingCellReference.RowIndex);
            Assert.AreEqual("D", result.EndingCellReference.ColumnName, true);
            Assert.AreEqual(4, result.EndingCellReference.ColumnIndex);
            Assert.AreEqual(target.RowIndex, result.EndingCellReference.RowIndex);
        }

        [TestMethod]
        public void ExtendColumnRange_NegativeArgument_CorrectRange()
        {
            var target = new SingleCellReference("D1");
            var result = target.ExtendColumnRange(-3) as RangeCellReference;

            Assert.AreEqual("A", result.StartingCellReference.ColumnName, true);
            Assert.AreEqual(1, result.StartingCellReference.ColumnIndex);
            Assert.AreEqual(target.RowIndex, result.StartingCellReference.RowIndex);
            Assert.AreEqual(target.ColumnName, result.EndingCellReference.ColumnName, true);
            Assert.AreEqual(target.ColumnIndex, result.EndingCellReference.ColumnIndex);
            Assert.AreEqual(target.RowIndex, result.EndingCellReference.RowIndex);
        }

        [TestMethod]
        public void ExtendColumnRange_NegativeArgumentPastExtent_CorrectRange()
        {
            var target = new SingleCellReference("D1");
            var result = target.ExtendColumnRange(-4) as RangeCellReference;

            Assert.AreEqual("A", result.StartingCellReference.ColumnName, true);
            Assert.AreEqual(1, result.StartingCellReference.ColumnIndex);
            Assert.AreEqual(target.RowIndex, result.StartingCellReference.RowIndex);
            Assert.AreEqual(target.ColumnName, result.EndingCellReference.ColumnName, true);
            Assert.AreEqual(target.ColumnIndex, result.EndingCellReference.ColumnIndex);
            Assert.AreEqual(target.RowIndex, result.EndingCellReference.RowIndex);
        }

        [TestMethod]
        public void ExtendColumnRange_ZeroArgument_ReturnsSingleCellReference()
        {
            var target = new SingleCellReference("B1");
            var result = target.ExtendColumnRange(0);

            Assert.IsInstanceOfType(result, typeof(SingleCellReference));
            Assert.IsTrue(target.Equals(result));
        }

        [TestMethod]
        public void ExtendColumnRange_NegativeArgumentBaseCell_ReturnsSingleCellReference()
        {
            var target = new SingleCellReference("A1");
            var result = target.ExtendColumnRange(-2);

            Assert.IsInstanceOfType(result, typeof(SingleCellReference));
            Assert.IsTrue(target.Equals(result));
        }

        [TestMethod]
        public void ExtendRowRange_PositiveArgument_CorrectRange()
        {
            var target = new SingleCellReference("A1");
            var result = target.ExtendRowRange(3) as RangeCellReference;

            Assert.AreEqual(target.ColumnName, result.StartingCellReference.ColumnName, true);
            Assert.AreEqual(target.ColumnIndex, result.StartingCellReference.ColumnIndex);
            Assert.AreEqual(target.RowIndex, result.StartingCellReference.RowIndex);
            Assert.AreEqual(target.ColumnName, result.EndingCellReference.ColumnName, true);
            Assert.AreEqual(target.ColumnIndex, result.EndingCellReference.ColumnIndex);
            Assert.AreEqual(4, result.EndingCellReference.RowIndex);
        }

        [TestMethod]
        public void ExtendRowRange_NegativeArgument_CorrectRange()
        {
            var target = new SingleCellReference("A4");
            var result = target.ExtendRowRange(-3) as RangeCellReference;

            Assert.AreEqual(target.ColumnName, result.StartingCellReference.ColumnName, true);
            Assert.AreEqual(target.ColumnIndex, result.StartingCellReference.ColumnIndex);
            Assert.AreEqual(1, result.StartingCellReference.RowIndex);
            Assert.AreEqual(target.ColumnName, result.EndingCellReference.ColumnName, true);
            Assert.AreEqual(target.ColumnIndex, result.EndingCellReference.ColumnIndex);
            Assert.AreEqual(target.RowIndex, result.EndingCellReference.RowIndex);
        }

        [TestMethod]
        public void ExtendRowRange_NegativeArgumentPastExtent_CorrectRange()
        {
            var target = new SingleCellReference("A4");
            var result = target.ExtendRowRange(-4) as RangeCellReference;

            Assert.AreEqual(target.ColumnName, result.StartingCellReference.ColumnName, true);
            Assert.AreEqual(target.ColumnIndex, result.StartingCellReference.ColumnIndex);
            Assert.AreEqual(1, result.StartingCellReference.RowIndex);
            Assert.AreEqual(target.ColumnName, result.EndingCellReference.ColumnName, true);
            Assert.AreEqual(target.ColumnIndex, result.EndingCellReference.ColumnIndex);
            Assert.AreEqual(target.RowIndex, result.EndingCellReference.RowIndex);
        }

        [TestMethod]
        public void ExtendRowRange_ZeroArgument_ReturnsSingleCellReference()
        {
            var target = new SingleCellReference("A1");
            var result = target.ExtendRowRange(0);

            Assert.IsInstanceOfType(result, typeof(SingleCellReference));
            Assert.IsTrue(target.Equals(result));
        }

        [TestMethod]
        public void ExtendRowRange_NegativeArgumentBaseCell_ReturnsSingleCellReference()
        {
            var target = new SingleCellReference("A1");
            var result = target.ExtendRowRange(-2);

            Assert.IsInstanceOfType(result, typeof(SingleCellReference));
            Assert.IsTrue(target.Equals(result));
        }

        #endregion
    }
}
