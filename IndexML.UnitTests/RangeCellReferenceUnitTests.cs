namespace IndexML.UnitTests
{
    using System;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Moq;

    /// <summary>
    /// Unit tests for the <see cref="RangeCellReference"/> class.
    /// </summary>
    [TestClass]
    public class RangeCellReferenceUnitTests : OpenXmlIndexerTest
    {
        #region Test Methods

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
        public void Constructor_ValidCellRange_CorrectProperties()
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

            Assert.IsFalse(target.ContainsOrSubsumes(mock.Object));
        }

        #endregion
    }
}
