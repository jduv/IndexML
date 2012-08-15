namespace IndexML.UnitTests
{
    using System;
    using System.Linq;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Unit tests for the <see cref="WorkbookIndexer"/> class.
    /// </summary>
    [TestClass]
    public class WorkbookIndexerUnitTests : OpenXmlIndexerTest
    {
        #region Test Methods

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Constructor_NullArgument_ThrowsException()
        {
            var target = new WorkbookIndexer(null);
        }

        [TestMethod]
        [DeploymentItem(EmptySheetPath, TestFilesDir)]
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
        [DeploymentItem(RandomDataSheetSpath, TestFilesDir)]
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
        [DeploymentItem(EmptyThreeSheetsPath, TestFilesDir)]
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
        [DeploymentItem(RandomDataThreeSheetSpath, TestFilesDir)]
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
