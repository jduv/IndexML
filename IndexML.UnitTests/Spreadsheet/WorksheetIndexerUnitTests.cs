namespace IndexML.UnitTests.Spreadsheet
{
    using System;
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using IndexML.Spreadsheet;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Unit tests for the <see cref="WorksheetIndexer"/> class.
    /// </summary>
    [TestClass]
    public class WorksheetIndexerUnitTests : SpreadsheetTest
    {
        #region Test Methods

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Constructor_NullArgument_ThrowsException()
        {
            var target = new WorksheetIndexer(null, null);
        }

        [TestMethod]
        [DeploymentItem(EmptySheetPath, TestFilesDir)]
        public void Constructor_EmptySheet_ValidState()
        {
            SafeExecuteTest(
                EmptySheetPath,
                (x) => x.WorkbookPart.WorksheetParts.First(),
                (wksPart) =>
                {
                    var target = new WorksheetIndexer(wksPart, null);

                    Assert.IsNotNull(target);
                    Assert.IsNotNull(target.SheetData);
                    Assert.IsTrue(string.IsNullOrEmpty(target.SheetName));
                });
        }

        [TestMethod]
        [DeploymentItem(ExactlyFiveRowsSheetPath, TestFilesDir)]
        public void Constructor_NonEmptySheet_ValidState()
        {
            SafeExecuteTest(
                ExactlyFiveRowsSheetPath,
                (x) => x.WorkbookPart.WorksheetParts.First(),
                (wksPart) =>
                {
                    var target = new WorksheetIndexer(wksPart, null);

                    Assert.IsNotNull(target);
                    Assert.IsNotNull(target.SheetData);
                    Assert.IsTrue(string.IsNullOrEmpty(target.SheetName));
                });
        }

        [TestMethod]
        [DeploymentItem(ExactlyFiveRowsSheetPath, TestFilesDir)]
        public void Constructor_NonEmptyWithSheetName_ValidState()
        {
            SafeExecuteTest(
                ExactlyFiveRowsSheetPath,
                (spreadsheet) =>
                {
                    var sheet = (Sheet)spreadsheet.WorkbookPart.Workbook.Sheets.First();
                    var worksheet = (WorksheetPart)spreadsheet.WorkbookPart.GetPartById(sheet.Id);
                    var target = new WorksheetIndexer(worksheet, sheet);

                    Assert.IsNotNull(target);
                    Assert.IsNotNull(target.SheetData);
                    Assert.IsFalse(string.IsNullOrEmpty(target.SheetName));
                    Assert.AreEqual(sheet.Name.ToString(), target.SheetName);
                });
        }

        #endregion
    }
}
