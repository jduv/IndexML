namespace IndexML.UnitTests.Spreadsheet
{
    using System;
    using System.Linq;
    using IndexML.Spreadsheet;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [DeploymentItem(@"IndexML.TestFiles\", @"IndexML.TestFiles\")]
    public class WorkbookIndexerUnitTests : SpreadsheetTest
    {
        #region Test Methods

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Constructor_NullArgument_ThrowsException()
        {
            var target = new WorkbookIndexer(null);
        }

        [TestMethod]
        public void Constructor_EmptySheet_HasOneSheet()
        {
            SafeExecuteTest(
                EmptySheetPath,
                (spreadsheet) =>
                {
                    var target = new WorkbookIndexer(spreadsheet.WorkbookPart);

                    Assert.IsNotNull(target);
                    Assert.IsTrue(target.Worksheets.Count() == 1);
                });
        }

        [TestMethod]
        public void Constructor_RandomData_ValidState()
        {
            SafeExecuteTest(
                RandomDataSheetSpath,
                (spreadsheet) =>
                {
                    var target = new WorkbookIndexer(spreadsheet.WorkbookPart);

                    Assert.IsNotNull(target);
                    Assert.IsTrue(target.Worksheets.Count() == 1);
                });
        }

        [TestMethod]
        public void Constructor_EmptyMultiSheets_HasThreeSheets()
        {
            SafeExecuteTest(
                EmptyThreeSheetsPath,
                (spreadsheet) =>
                {
                    var target = new WorkbookIndexer(spreadsheet.WorkbookPart);

                    Assert.IsNotNull(target);
                    Assert.IsTrue(target.Worksheets.Count() == 3);
                });
        }

        [TestMethod]
        public void Constructor_RandomDataMultiSheets_HasThreeSheets()
        {
            SafeExecuteTest(
                RandomDataThreeSheetSpath,
                (spreadsheet) =>
                {
                    var target = new WorkbookIndexer(spreadsheet.WorkbookPart);

                    Assert.IsNotNull(target);
                    Assert.IsTrue(target.Worksheets.Count() == 3);
                });
        }

        #endregion
    }
}
