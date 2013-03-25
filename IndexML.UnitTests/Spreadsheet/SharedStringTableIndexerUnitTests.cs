namespace IndexML.UnitTests.Spreadsheet
{
    using System;
    using System.Collections.Generic;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using IndexML.Spreadsheet;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    public class SharedStringTableIndexerUnitTests : SpreadsheetTest
    {
        #region Test Methods

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Constructor_NullArgumentPart_ThrowsException()
        {
            var target = new SharedStringTableIndexer((SharedStringTablePart)null);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Constructor_NullArgumentSharedStringTable_ThrowsException()
        {
            var target = new SharedStringTableIndexer((SharedStringTable)null);
        }

        [TestMethod]
        [DeploymentItem(EmptySheetPath, TestFilesDir)]
        public void Constructor_NullStringTableOnPart_CreatesEmptyIndexer()
        {
            SafeExecuteTest(
                EmptySheetPath,
                (spreadsheet) =>
                {
                    var sstPart = spreadsheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
                    var target = new SharedStringTableIndexer(sstPart);

                    Assert.IsNotNull(target);
                    Assert.AreEqual(0, target.UniqueCount);
                    Assert.IsNotNull(spreadsheet.WorkbookPart.SharedStringTablePart.SharedStringTable); // constructor is side affecting
                });
        }

        [TestMethod]
        [DeploymentItem(RandomDataSheetSpath, TestFilesDir)]
        public void Constructor_SharedStringTablePartWithData_NonEmptyIndexer()
        {
            SafeExecuteTest(
                RandomDataSheetSpath,
                (spreadsheet) =>
                {
                    var target = new SharedStringTableIndexer(spreadsheet.WorkbookPart.SharedStringTablePart);

                    Assert.IsNotNull(target);
                    Assert.AreNotEqual(0, target.UniqueCount);
                });
        }

        [TestMethod]
        [DeploymentItem(RandomDataSheetSpath, TestFilesDir)]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Add_NullString_ThrowsException()
        {
            SafeExecuteTest(
                RandomDataSheetSpath,
                (spreadsheet) =>
                {
                    var target = new SharedStringTableIndexer(spreadsheet.WorkbookPart.SharedStringTablePart);
                    target.Add(null);
                });
        }

        [TestMethod]
        [DeploymentItem(RandomDataSheetSpath, TestFilesDir)]
        public void Add_NonEmptyStringTwice_AddedAtReturnedIndexNoDuplicates()
        {
            SafeExecuteTest(
                RandomDataSheetSpath,
                (spreadsheet) =>
                {
                    var target = new SharedStringTableIndexer(spreadsheet.WorkbookPart.SharedStringTablePart);
                    var item = Guid.NewGuid().ToString();
                    var oldCount = target.UniqueCount;

                    // add it twice
                    var index1 = target.Add(item);
                    var index2 = target.Add(item);

                    Assert.AreEqual(index1, index2);
                    Assert.AreEqual(oldCount + 1, target.UniqueCount); // should only increment once.
                });
        }

        [TestMethod]
        [DeploymentItem(RandomDataSheetSpath, TestFilesDir)]
        public void Add_NonEmptyValue_AddedAtReturnedIndex()
        {
            SafeExecuteTest(
                RandomDataSheetSpath,
                (spreadsheet) =>
                {
                    var target = new SharedStringTableIndexer(spreadsheet.WorkbookPart.SharedStringTablePart);
                    var item = Guid.NewGuid().ToString();
                    var oldCount = target.UniqueCount;
                    var index = target.Add(item);

                    Assert.AreEqual(oldCount + 1, target.UniqueCount);
                    Assert.AreEqual(index, target[item]);
                    Assert.AreEqual(item, target[index]);
                });
        }

        [TestMethod]
        [DeploymentItem(RandomDataSheetSpath, TestFilesDir)]
        [ExpectedException(typeof(ArgumentNullException))]
        public void AddAll_NullArgument_ThrowsException()
        {
            SafeExecuteTest(
                RandomDataSheetSpath,
                (spreadsheet) =>
                {
                    var target = new SharedStringTableIndexer(spreadsheet.WorkbookPart.SharedStringTablePart);
                    var oldCount = target.UniqueCount;

                    target.AddAll(null);

                    Assert.AreEqual(oldCount, target.UniqueCount);
                });
        }

        [TestMethod]
        [DeploymentItem(RandomDataSheetSpath, TestFilesDir)]
        public void AddAll_EmptyList_NoChange()
        {
            SafeExecuteTest(
                RandomDataSheetSpath,
                (spreadsheet) =>
                {
                    var target = new SharedStringTableIndexer(spreadsheet.WorkbookPart.SharedStringTablePart);
                    var oldCount = target.UniqueCount;

                    target.AddAll(new string[0]);

                    Assert.AreEqual(oldCount, target.UniqueCount);
                });
        }

        [TestMethod]
        [DeploymentItem(RandomDataSheetSpath, TestFilesDir)]
        public void AddAll_ValidList_AddedAtReturnedIndices()
        {
            SafeExecuteTest(
                RandomDataSheetSpath,
                (spreadsheet) =>
                {
                    var toAdd = new string[] { Guid.NewGuid().ToString(), Guid.NewGuid().ToString(), Guid.NewGuid().ToString() };
                    var target = new SharedStringTableIndexer(spreadsheet.WorkbookPart.SharedStringTablePart);
                    var oldCount = target.UniqueCount;

                    var indices = target.AddAll(toAdd);

                    Assert.AreEqual(oldCount + toAdd.Length, target.UniqueCount);

                    int i = 0;
                    foreach (var index in indices)
                    {
                        Assert.AreEqual(toAdd[i], target[index]);
                        i++;
                    }

                    foreach (var item in toAdd)
                    {
                        Assert.IsNotNull(target[item]);
                    }
                });
        }

        [TestMethod]
        [DeploymentItem(RandomDataSheetSpath, TestFilesDir)]
        public void Contains_NonExistantStringAndIndex_ReturnsFalse()
        {
            SafeExecuteTest(
                RandomDataSheetSpath,
                (spreadsheet) =>
                {
                    var target = new SharedStringTableIndexer(spreadsheet.WorkbookPart.SharedStringTablePart);
                    Assert.IsFalse(target.Contains(Guid.NewGuid().ToString()));
                    Assert.IsFalse(target.Contains(target.UniqueCount + 1));
                });
        }

        [TestMethod]
        [DeploymentItem(RandomDataSheetSpath, TestFilesDir)]
        public void Contains_ExistingStringAndIndex_ReturnsFalse()
        {
            SafeExecuteTest(
                RandomDataSheetSpath,
                (spreadsheet) =>
                {
                    var target = new SharedStringTableIndexer(spreadsheet.WorkbookPart.SharedStringTablePart);
                    var item = Guid.NewGuid().ToString();
                    var index = target.Add(item);

                    Assert.IsTrue(target.Contains(item));
                    Assert.IsTrue(target.Contains(index));
                });
        }

        [TestMethod]
        [DeploymentItem(RandomDataSheetSpath, TestFilesDir)]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Indexer_NullString_ThrowsException()
        {
            SafeExecuteTest(
                RandomDataSheetSpath,
                (spreadsheet) =>
                {
                    var target = new SharedStringTableIndexer(spreadsheet.WorkbookPart.SharedStringTablePart);
                    var value = target[null];
                });
        }

        [TestMethod]
        [DeploymentItem(RandomDataSheetSpath, TestFilesDir)]
        [ExpectedException(typeof(KeyNotFoundException))]
        public void Indexer_NonExistingString_ThrowsException()
        {
            SafeExecuteTest(
                RandomDataSheetSpath,
                (spreadsheet) =>
                {
                    var target = new SharedStringTableIndexer(spreadsheet.WorkbookPart.SharedStringTablePart);
                    var value = target[Guid.NewGuid().ToString()];
                });
        }

        [TestMethod]
        [DeploymentItem(RandomDataSheetSpath, TestFilesDir)]
        [ExpectedException(typeof(KeyNotFoundException))]
        public void Indexer_NegativeIndex_ThrowsException()
        {
            SafeExecuteTest(
                RandomDataSheetSpath,
                (spreadsheet) =>
                {
                    var target = new SharedStringTableIndexer(spreadsheet.WorkbookPart.SharedStringTablePart);
                    var value = target[-1];
                });
        }

        #endregion
    }
}
