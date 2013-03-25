namespace IndexML.UnitTests.Spreadsheet
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using DocumentFormat.OpenXml.Spreadsheet;
    using IndexML.Spreadsheet;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    public class DataValidationIndexerUnitTests : SpreadsheetTest
    {
        #region Fields & Constants

        private static readonly Cell A1 = new Cell() { CellReference = "A1" };

        private static readonly Cell A2 = new Cell() { CellReference = "A2" };

        private static readonly Cell A3 = new Cell() { CellReference = "A3" };

        private static readonly Cell B2 = new Cell() { CellReference = "B2" };

        private static readonly Cell B4 = new Cell() { CellReference = "B4" };

        private static readonly Cell C6 = new Cell() { CellReference = "C6" };

        #endregion

        #region Test Methods

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Constructor_NullArgument_ThrowsException()
        {
            var target = new DataValidationIndexer(null);
        }

        [TestMethod]
        [DeploymentItem(StaggeredValidationsSheetPath, TestFilesDir)]
        public void Constructor_StaggeredSingleCellValidations_CorrectCellReferences()
        {
            SafeExecuteTest(
                StaggeredValidationsSheetPath,
                (validations) =>
                {
                    var target = new DataValidationIndexer(validations.First());
                    var refs = target.CellReferences.ToList();

                    Assert.AreEqual(3, refs.Count);
                    Assert.AreEqual("A2", refs[0].Value, true);
                    Assert.AreEqual("B4", refs[1].Value, true);
                    Assert.AreEqual("C6", refs[2].Value, true);
                });
        }

        [TestMethod]
        [DeploymentItem(AllValidationsSheetPath, TestFilesDir)]
        public void Add_NullCell_Ignored()
        {
            SafeExecuteTest(
                AllValidationsSheetPath,
                (validations) =>
                {
                });
        }

        [TestMethod]
        [DeploymentItem(AllValidationsSheetPath, TestFilesDir)]
        public void Add_NullCellReference_Ignored()
        {
            SafeExecuteTest(
                AllValidationsSheetPath,
                (validations) =>
                {
                    var malformedCell = new Cell()
                    {
                        CellReference = null
                    };

                    var target = new DataValidationIndexer(validations.First());
                    target.Add(malformedCell);
                });
        }

        [TestMethod]
        [DeploymentItem(OneValidationA2SheetPath, TestFilesDir)]
        public void Add_EmptyCellreference_Ignored()
        {
            SafeExecuteTest(
                AllValidationsSheetPath,
                (validations) =>
                {
                    var malformedCell = new Cell()
                    {
                        CellReference = string.Empty
                    };

                    var target = new DataValidationIndexer(validations.First());
                    target.Add(malformedCell);
                });
        }

        [TestMethod]
        [DeploymentItem(OneValidationA2SheetPath, TestFilesDir)]
        [ExpectedException(typeof(ArgumentException))]
        public void Add_MalformedCellReference_ExceptionThrown()
        {
            SafeExecuteTest(
                OneValidationA2SheetPath,
                (validations) =>
                {
                    var malformedCell = new Cell()
                    {
                        CellReference = Guid.NewGuid().ToString()
                    };

                    var target = new DataValidationIndexer(validations.First());
                    target.Add(malformedCell);
                });
        }

        [TestMethod]
        [DeploymentItem(OneValidationA2SheetPath, TestFilesDir)]
        public void Add_ExistingSingleCellReference_NotAdded()
        {
            SafeExecuteTest(
                OneValidationA2SheetPath,
                (validations) =>
                {
                    var target = new DataValidationIndexer(validations.First());
                    var targetBeforeCount = target.ReferenceCount;
                    target.Add(A2);

                    Assert.AreEqual(targetBeforeCount, target.ReferenceCount);
                });
        }

        [TestMethod]
        [DeploymentItem(RowValidationsSheetPath, TestFilesDir)]
        public void Add_NewSingleCellToRowRangeValidator_ReferenceListIncludesCell()
        {
            SafeExecuteTest(
                RowValidationsSheetPath,
                (validations) =>
                {
                    var target = new DataValidationIndexer(validations.First());
                    var targetBeforeCount = target.ReferenceCount;
                    target.Add(A2);

                    Assert.AreEqual(targetBeforeCount + 1, target.ReferenceCount);
                });
        }

        [TestMethod]
        [DeploymentItem(ColumnValidationsSheetPath, TestFilesDir)]
        public void Add_NewSingleCellToColumnRangeValidator_ReferenceListIncludesCell()
        {
            SafeExecuteTest(
                ColumnValidationsSheetPath,
                (validations) =>
                {
                    var target = new DataValidationIndexer(validations.First());
                    var targetBeforeCount = target.ReferenceCount;
                    target.Add(B2);

                    Assert.AreEqual(targetBeforeCount + 1, target.ReferenceCount);
                });
        }

        [TestMethod]
        [DeploymentItem(ColumnValidationsSheetPath, TestFilesDir)]
        public void Add_SingleCellReferenceCollidesWithRangeValidator_NotAdded()
        {
            SafeExecuteTest(
                ColumnValidationsSheetPath,
                (validations) =>
                {
                    var target = new DataValidationIndexer(validations.First());
                    var targetBeforeCount = target.ReferenceCount;
                    target.Add(A3);

                    Assert.AreEqual(targetBeforeCount, target.ReferenceCount);
                });
        }

        [TestMethod]
        [DeploymentItem(StaggeredValidationsSheetPath, TestFilesDir)]
        public void Clear_StaggeredValidations_NoItems()
        {
            SafeExecuteTest(
                StaggeredValidationsSheetPath,
                (validations) =>
                {
                    var target = new DataValidationIndexer(validations.First());
                    target.Clear();

                    Assert.AreEqual(0, target.ReferenceCount);
                    Assert.AreEqual(0, target.CellReferences.Count());
                });
        }

        ////[TestMethod]
        ////[DeploymentItem(ColumnValidationsSheetPath, TestFilesDir)]
        ////public void Remove_SingleCellReferenceNonExisting_NotRemoved()
        ////{
        ////    var aa1 = new Cell()
        ////    {
        ////        CellReference = "AA1"
        ////    };

        ////    SafeExecuteTest(
        ////        ColumnValidationsSheetPath,
        ////        (validations) =>
        ////        {
        ////            var target = new DataValidationIndexer(validations.First());
        ////            target.Remove(aa1);

        ////            // Only the range should exist
        ////            Assert.AreEqual(1, target.ReferenceCount);
        ////            Assert.AreEqual(1, target.CellReferences.Count());
        ////        });
        ////}

        ////[TestMethod]
        ////[DeploymentItem(StaggeredValidationsSheetPath, TestFilesDir)]
        ////public void Remove_SingleCellReferenceExisting_Removed()
        ////{
        ////    var a2 = new Cell()
        ////    {
        ////        CellReference = "A2"
        ////    };

        ////    SafeExecuteTest(
        ////        StaggeredValidationsSheetPath,
        ////        (validations) =>
        ////        {
        ////            var target = new DataValidationIndexer(validations.First());                    
        ////            target.Remove(a2);

        ////            Assert.AreEqual(2, target.ReferenceCount);
        ////            Assert.AreEqual(2, target.CellReferences.Count());
        ////        });
        ////}

        ////[TestMethod]
        ////[DeploymentItem(ColumnValidationsSheetPath, TestFilesDir)]
        ////public void Remove_SingleCellReferenceExistingRangeCollision_RangeSplit()
        ////{
        ////    var b2 = new Cell()
        ////    {
        ////        CellReference = "B1"
        ////    };

        ////    SafeExecuteTest(
        ////        ColumnValidationsSheetPath,
        ////        (validations) =>
        ////        {
        ////            var target = new DataValidationIndexer(validations.First());
        ////            target.Remove(b2);

        ////            // One range will be split into two single references
        ////            Assert.AreEqual(2, target.ReferenceCount);
        ////            Assert.AreEqual(2, target.CellReferences.Count());
        ////        });
        ////}

        ////[TestMethod]
        ////[DeploymentItem(ColumnValidationsSheetPath, TestFilesDir)]
        ////public void Remove_SingleCellReferenceExistingRangeNoCollision_NoChange()
        ////{
        ////    var aa1 = new Cell()
        ////    {
        ////        CellReference = "AA1"
        ////    };

        ////    SafeExecuteTest(
        ////        ColumnValidationsSheetPath,
        ////        (validations) =>
        ////        {
        ////            var target = new DataValidationIndexer(validations.First());
        ////            target.Remove(aa1);

        ////            // only one reference, the range
        ////            Assert.AreEqual(1, target.ReferenceCount);
        ////            Assert.AreEqual(1, target.CellReferences.Count());
        ////        });
        ////}

        [TestMethod]
        [DeploymentItem(StaggeredValidationsSheetPath, TestFilesDir)]
        public void Contains_ExistingSingleCellRefs_ReturnsTrue()
        {
            SafeExecuteTest(
                StaggeredValidationsSheetPath,
                (validations) =>
                {
                    var target = new DataValidationIndexer(validations.First());
                    Assert.IsTrue(target.Contains(A2));
                    Assert.IsTrue(target.Contains(B4));
                    Assert.IsTrue(target.Contains(C6));
                });
        }

        [TestMethod]
        [DeploymentItem(StaggeredValidationsSheetPath, TestFilesDir)]
        public void Contains_NonExistantSingleCellRef_ReturnsFalse()
        {
            SafeExecuteTest(
                StaggeredValidationsSheetPath,
                (validations) =>
                {
                    var target = new DataValidationIndexer(validations.First());
                    Assert.IsFalse(target.Contains(A1));
                });
        }

        [TestMethod]
        [DeploymentItem(RowValidationsSheetPath, TestFilesDir)]
        public void Contains_Range_Collision_ReturnsTrue()
        {
            SafeExecuteTest(
                RowValidationsSheetPath,
                (validations) =>
                {
                    var target = new DataValidationIndexer(validations.First());
                    Assert.IsFalse(target.Contains(A2));
                });
        }

        [TestMethod]
        [DeploymentItem(RowValidationsSheetPath, TestFilesDir)]
        public void Contains_RangeNoCollision_ReturnsFalse()
        {
            SafeExecuteTest(
                RowValidationsSheetPath,
                (validations) =>
                {
                    var target = new DataValidationIndexer(validations.First());
                    Assert.IsFalse(target.Contains(B2));
                });
        }

        #endregion

        #region Private Methods

        private static void SafeExecuteTest(string spreadsheetPath, Action<IEnumerable<DataValidation>> test)
        {
            SpreadsheetTest.SafeExecuteTest<IEnumerable<DataValidation>>(
                spreadsheetPath,
                x => x.WorkbookPart.WorksheetParts.SelectMany(w => w.Worksheet.Descendants<DataValidation>()),
                test);
        }

        #endregion
    }
}
