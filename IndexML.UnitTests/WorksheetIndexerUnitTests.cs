namespace IndexML.UnitTests
{
    using System;
    using System.Linq;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Unit tests for the <see cref="WorksheetIndexer"/> class.
    /// </summary>
    [TestClass]
    public class WorksheetIndexerUnitTests : OpenXmlIndexerTest
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
                    Assert.IsTrue(target.SheetData.Count > 0);
                });           
        }        

        #endregion
    }
}
