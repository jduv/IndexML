namespace IndexML.UnitTests.Spreadsheet
{
    using System;
    using System.Linq;
    using DocumentFormat.OpenXml.Spreadsheet;
    using IndexML.Spreadsheet;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [DeploymentItem(@"IndexML.TestFiles\", @"IndexML.TestFiles\")]
    public class SheetDataIndexerUnitTests : SpreadsheetTest
    {
        #region Test Methods

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Constructor_NullArgument_ExceptionThrown()
        {
            var target = new ArrayBasedSheetDataIndexer(null);
        }

        [TestMethod]
        public void Constructor_EmptySheetData_ValidEmptyState()
        {
            SafeExecuteTest(
                EmptySheetPath,                
                (sheetData) =>
                {
                    var target = new ArrayBasedSheetDataIndexer(sheetData);

                    Assert.IsNotNull(target);
                    Assert.IsTrue(target.IsEmpty);
                    Assert.AreEqual(0, target.Count);
                });
        }

        [TestMethod]
        public void Constructor_ValidSheetData_ValidState()
        {
            SafeExecuteTest(
                RandomDataSheetSpath,                
                (sheetData) =>
                {
                    var target = new ArrayBasedSheetDataIndexer(sheetData);

                    Assert.IsNotNull(target);
                    Assert.IsFalse(target.IsEmpty);

                    // These magic numbers can only be determined from the test spreadsheet.
                    Assert.AreEqual(10, target.Count);
                    Assert.AreEqual(14, target.MaxRowIndex);
                    ValidateRowSequence(target);
                });
        }

        [TestMethod]
        public void Constructor_MaxExtents_ValidState()
        {
            SafeExecuteTest(
                MaxExtentsSheetPath,                
                (sheetData) =>
                {
                    var target = new ArrayBasedSheetDataIndexer(sheetData);

                    Assert.IsNotNull(target);
                    Assert.IsFalse(target.IsEmpty);
                    Assert.AreEqual(2, target.Count);
                    Assert.AreEqual(ArrayBasedSheetDataIndexer.Capacity, target.MaxRowIndex);
                    ValidateRowSequence(target);
                });
        }

        [TestMethod]
        [ExpectedException(typeof(IndexOutOfRangeException))]
        public void AppendRow_MaxExtents_ThrowsException()
        {
            SafeExecuteTest(
                MaxExtentsSheetPath,                
                (sheetData) =>
                {
                    var target = new ArrayBasedSheetDataIndexer(sheetData);
                    target.AppendRow(new Row());
                });
        }

        [TestMethod]
        public void AppendRow_EmptySpreadsheet_IncreasesCountByOne()
        {
            SafeExecuteTest(
                EmptySheetPath,                
                (sheetData) =>
                {
                    var target = new ArrayBasedSheetDataIndexer(sheetData);
                    target.AppendRow(new Row());

                    Assert.IsFalse(target.IsEmpty);
                    Assert.AreEqual(1, target.Count);
                    Assert.AreEqual(1, target.MaxRowIndex); // Row indices are one based.
                    ValidateRowSequence(target);
                });
        }

        [TestMethod]
        [ExpectedException(typeof(IndexOutOfRangeException))]
        public void RemoveRow_NegativeIndex_ThrowsException()
        {
            SafeExecuteTest(
                EmptySheetPath,                
                (sheetData) =>
                {
                    var target = new ArrayBasedSheetDataIndexer(sheetData);
                    target.RemoveRow(-1);
                });
        }

        [TestMethod]
        [ExpectedException(typeof(IndexOutOfRangeException))]
        public void RemoveRow_OverCapacityIndex_ThrowsException()
        {
            SafeExecuteTest(
                EmptySheetPath,                
                (sheetData) =>
                {
                    var target = new ArrayBasedSheetDataIndexer(sheetData);
                    target.RemoveRow(ArrayBasedSheetDataIndexer.Capacity + 2); // Row indices are one based, hence the +2 is needed
                });
        }

        [TestMethod]   
        public void RemoveRow_EmptySpreadsheet_DoesNothing()
        {
            SafeExecuteTest(
                EmptySheetPath,                
                (sheetData) =>
                {
                    var target = new ArrayBasedSheetDataIndexer(sheetData);
                    target.RemoveRow(0);
                });
        }

        [TestMethod]
        public void RemoveRow_NonExistentRow_ReturnsFalse()
        {
            SafeExecuteTest(
                ExactlyFiveRowsSheetPath,                
                (sheetData) =>
                {
                    var target = new ArrayBasedSheetDataIndexer(sheetData);
                    Assert.IsFalse(target.RemoveRow(6));
                    ValidateRowSequence(target);
                });
        }

        [TestMethod]
        public void RemoveRow_MaxRowNoShift_DecreasesCountByOne()
        {
            SafeExecuteTest(
                ExactlyFiveRowsSheetPath,                
                (sheetData) =>
                {
                    var target = new ArrayBasedSheetDataIndexer(sheetData);

                    var oldCount = target.Count;
                    var oldMaxRowIndex = target.MaxRowIndex;

                    Assert.IsTrue(target.RemoveRow(target.MaxRowIndex));
                    Assert.AreEqual(oldCount, target.Count + 1);
                    Assert.AreNotEqual(oldMaxRowIndex, target.MaxRowIndex); // Might not necessarily be minus one.
                    ValidateRowSequence(target);
                });
        }

        [TestMethod]
        public void RemoveRow_SequentialMiddleRowNoShiftUp_DecreasesCountAndMaxRowByOne()
        {
            SafeExecuteTest(
                ExactlyFiveRowsSheetPath,                
                (sheetData) =>
                {
                    var target = new ArrayBasedSheetDataIndexer(sheetData);

                    var oldCount = target.Count;
                    var oldMaxRowIndex = target.MaxRowIndex;

                    Assert.IsTrue(target.RemoveRow(target.Rows.First().RowIndex));
                    Assert.AreEqual(oldCount, target.Count + 1);
                    Assert.AreEqual(oldMaxRowIndex, target.MaxRowIndex);
                    ValidateRowSequence(target);
                });
        }

        [TestMethod]
        public void RemoveRow_MaxRowShift_DecreasesCountByOne()
        {
            SafeExecuteTest(
                ExactlyFiveRowsSheetPath,                
                (sheetData) =>
                {
                    var target = new ArrayBasedSheetDataIndexer(sheetData);

                    var oldCount = target.Count;
                    var oldMaxRowIndex = target.MaxRowIndex;

                    Assert.IsTrue(target.RemoveRow(target.MaxRowIndex, true));
                    Assert.AreEqual(oldCount - 1, target.Count);
                    Assert.AreNotEqual(oldMaxRowIndex, target.MaxRowIndex); // Might not necessarily be minus one.
                    ValidateRowSequence(target);
                });
        }

        [TestMethod]
        public void RemoveRow_SequentialMiddleRowShiftUp_DecreasesCountAndMaxRowByOne()
        {
            SafeExecuteTest(
                ExactlyFiveRowsSheetPath,                
                (sheetData) =>
                {
                    var target = new ArrayBasedSheetDataIndexer(sheetData);

                    var oldCount = target.Count;
                    var oldMaxRowIndex = target.MaxRowIndex;

                    Assert.IsTrue(target.RemoveRow(target.Rows.First().RowIndex, true));
                    Assert.AreEqual(oldCount - 1, target.Count);
                    Assert.AreEqual(oldMaxRowIndex - 1, target.MaxRowIndex);
                    ValidateRowSequence(target);
                });
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void InsertRow_NullRowArgument_ThrowsException()
        {
            SafeExecuteTest(
                EmptySheetPath,                
                (sheetData) =>
                {                    
                    var target = new ArrayBasedSheetDataIndexer(sheetData);
                    target.InsertRow(null, 1);
                });
        }

        [TestMethod]
        [ExpectedException(typeof(IndexOutOfRangeException))]
        public void InsertRow_NegativeIndex_ThrowsException()
        {
            SafeExecuteTest(
                ExactlyFiveRowsSheetPath,                
                (sheetData) =>
                {
                    var target = new ArrayBasedSheetDataIndexer(sheetData);
                    target.InsertRow(new Row(), -1);
                });
        }

        [TestMethod]
        [ExpectedException(typeof(IndexOutOfRangeException))]
        public void InsertRow_IndexOverCapacity_ThrowsException()
        {
            SafeExecuteTest(
                ExactlyFiveRowsSheetPath,                
                (sheetData) =>
                {
                    var target = new ArrayBasedSheetDataIndexer(sheetData);
                    target.InsertRow(new Row(), ArrayBasedSheetDataIndexer.Capacity + 1);
                });
        }

        [TestMethod]
        public void InsertRow_SequentialMiddleRowNoShift_DoesNotIncreaseCountOrMaxRow()
        {
            SafeExecuteTest(
                ExactlyFiveRowsSheetPath,
                (sheetData) =>
                {
                    var target = new ArrayBasedSheetDataIndexer(sheetData);
                    var halfway = target.Rows.ToList().Count / 2;

                    var oldCount = target.Count;
                    var oldMaxRow = target.MaxRowIndex;

                    target.InsertRow(new Row(), halfway);
                    Assert.IsFalse(target.IsEmpty);
                    Assert.AreEqual(oldCount, target.Count);
                    Assert.AreEqual(oldMaxRow, target.MaxRowIndex);
                    ValidateRowSequence(target);
                });
        }

        [TestMethod]
        public void InsertRow_SequentialMiddleRowShift_IncreasesCountAndMaxRowByOne()
        {
            SafeExecuteTest(
                ExactlyFiveRowsSheetPath,                
                (sheetData) =>
                {
                    var target = new ArrayBasedSheetDataIndexer(sheetData);
                    var halfway = (target.Rows.ToList().Count / 2) + 1;

                    var oldCount = target.Count;
                    var oldMaxRow = target.MaxRowIndex;

                    target.InsertRow(new Row(), halfway, true);
                    Assert.IsFalse(target.IsEmpty);
                    Assert.AreEqual(oldCount + 1, target.Count);
                    Assert.AreEqual(oldMaxRow + 1, target.MaxRowIndex);
                    ValidateRowSequence(target);
                });
        }

        [TestMethod]
        public void InsertRow_SequentialMaxRowNoShift_IncreasesCountByOne()
        {
            SafeExecuteTest(
                ExactlyFiveRowsSheetPath,                
                (sheetData) =>
                {
                    var target = new ArrayBasedSheetDataIndexer(sheetData);

                    var oldCount = target.Count;
                    var oldMaxRow = target.MaxRowIndex;

                    target.InsertRow(new Row(), target.MaxRowIndex);
                    Assert.IsFalse(target.IsEmpty);
                    Assert.AreEqual(oldCount, target.Count);
                    Assert.AreEqual(oldMaxRow, target.MaxRowIndex);
                    ValidateRowSequence(target);
                });
        }

        [TestMethod]
        public void InsertRow_SequentialMaxRowShift_IncreasesCountAndMaxRowByOne()
        {
            SafeExecuteTest(
                ExactlyFiveRowsSheetPath,                
                (sheetData) =>
                {
                    var target = new ArrayBasedSheetDataIndexer(sheetData);

                    var oldCount = target.Count;
                    var oldMaxRow = target.MaxRowIndex;

                    target.InsertRow(new Row(), target.MaxRowIndex, true);
                    Assert.IsFalse(target.IsEmpty);
                    Assert.AreEqual(oldCount + 1, target.Count);
                    Assert.AreEqual(oldMaxRow + 1, target.MaxRowIndex);
                    ValidateRowSequence(target);
                });
        }

        [TestMethod]
        public void InsertRow_NonExistingIndexNoShift_IncreasesCountButNotMaxRow()
        {
            SafeExecuteTest(
                FiveEvenRowsSheetPath,
                (sheetData) =>
                {
                    var target = new ArrayBasedSheetDataIndexer(sheetData);

                    var oldCount = target.Count;
                    var oldMaxRow = target.MaxRowIndex;

                    target.InsertRow(new Row(), 1); // insert at an odd index, test file has only even rows
                    Assert.IsFalse(target.IsEmpty);
                    Assert.AreEqual(oldCount + 1, target.Count);
                    Assert.AreEqual(oldMaxRow, target.MaxRowIndex);
                    ValidateRowSequence(target);
                });
        }

        [TestMethod]
        public void InsertRow_NonExistingIndexShift_IncreasesCountAndMaxRow()
        {
            SafeExecuteTest(
                FiveEvenRowsSheetPath,                
                (sheetData) =>
                {
                    var target = new ArrayBasedSheetDataIndexer(sheetData);

                    var oldCount = target.Count;
                    var oldMaxRow = target.MaxRowIndex;

                    target.InsertRow(new Row(), 1, true); // insert at an odd index, test file has only even rows
                    Assert.IsFalse(target.IsEmpty);
                    Assert.AreEqual(oldCount + 1, target.Count);
                    Assert.AreEqual(oldMaxRow + 1, target.MaxRowIndex);
                    ValidateRowSequence(target);
                });
        }

        [TestMethod]
        public void RowsProperty_ExactlyFiveRows_NoNulls()
        {
            SafeExecuteTest(
                ExactlyFiveRowsSheetPath,                
                (sheetData) =>
                {
                    var target = new ArrayBasedSheetDataIndexer(sheetData);
                    var rows = target.Rows.ToList();

                    Assert.AreEqual(5, rows.Count);
                    Assert.AreEqual(target.Count, rows.Count);
                    Assert.IsTrue(rows.TrueForAll(x => x != null));
                    ValidateRowSequence(target);
                });
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void MaxRowIndexProperty_EmptySheetData_ThrowsException()
        {
            SafeExecuteTest(
                EmptySheetPath,                
                (sheetData) =>
                {
                    var target = new ArrayBasedSheetDataIndexer(sheetData);

                    Assert.IsNotNull(target);
                    Assert.IsTrue(target.IsEmpty);
                    Assert.AreEqual(0, target.Count);
                    var index = target.MaxRowIndex; // should throw here
                });
        }

        [TestMethod]
        public void ImplicitCast_ValidIndexer_SameReference()
        {
            SafeExecuteTest(
                EmptySheetPath,
                (sheetData) =>
                {
                    var indexer = new ArrayBasedSheetDataIndexer(sheetData);
                    var target = (SheetData)indexer;
                    Assert.AreSame(sheetData, target);
                });
        }

        [TestMethod]
        public void ImplicitCast_Null_IsNull()
        {
            ArrayBasedSheetDataIndexer indexer = null;
            var target = (SheetData)indexer;
            Assert.IsNull(target);
        }

        #endregion

        #region Private Methods

        private static void SafeExecuteTest(string spreadsheetPath, Action<SheetData> testToPerform)
        {
            SafeExecuteTest(
                spreadsheetPath,
                (x) => x.WorkbookPart.WorksheetParts.First().Worksheet.Descendants<SheetData>().First(),
                testToPerform);
        }

        #endregion
    }
}
