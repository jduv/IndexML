namespace IndexML.UnitTests
{
    using System;
    using System.IO;
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Unit tests for the <see cref="SpreadsheetIndexer"/> class.
    /// </summary>
    [TestClass]
    public class SpreadsheetIndexerUnitTests : OpenXmlIndexerTest
    {
        #region Test Methods

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Constructor_NullArgument_ThrowsException()
        {
            var target = new SpreadsheetIndexer((byte[])null);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Constructor_NullStream_ThrowsException()
        {
            var target = new SpreadsheetIndexer((Stream)null);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void Constructor_EmptyStream_ThrowsException()
        {
            var target = new SpreadsheetIndexer(new MemoryStream());
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void Constructor_UnreadableStream_ThrowsException()
        {
            var stream = new MemoryStream();
            stream.Dispose(); // makes the stream unreadable

            var target = new SpreadsheetIndexer(stream);
        }

        [TestMethod]
        [DeploymentItem(EmptySheetPath, TestFilesDir)]
        public void Constructor_EmptySheet_ValidState()
        {
            var spreadsheetBytes = LoadTestSpreadSheetBytes(EmptySheetPath);
            if (spreadsheetBytes != null)
            {
                using (var target = new SpreadsheetIndexer(spreadsheetBytes))
                {
                    Assert.IsNotNull(target);
                    Assert.IsNotNull(target.Workbook);
                    Assert.IsTrue(target.Workbook.Worksheets.Count() == 1);
                }
            }
        }

        [TestMethod]
        [DeploymentItem(RandomDataSheetSpath, TestFilesDir)]
        public void Constructor_RandomData_ValidState()
        {
            var spreadsheetBytes = LoadTestSpreadSheetBytes(RandomDataSheetSpath);
            if (spreadsheetBytes != null)
            {
                using (var target = new SpreadsheetIndexer(spreadsheetBytes))
                {
                    Assert.IsNotNull(target);
                    Assert.IsNotNull(target.Workbook);
                    Assert.IsTrue(target.Workbook.Worksheets.Count() == 1);
                    Assert.AreEqual(spreadsheetBytes.Length, target.GetBytes().Length);
                }
            }
        }

        [TestMethod]
        [DeploymentItem(EmptyThreeSheetsPath, TestFilesDir)]
        public void Constructor_EmptyMultiSheets_ValidState()
        {
            var spreadsheetBytes = LoadTestSpreadSheetBytes(EmptyThreeSheetsPath);
            if (spreadsheetBytes != null)
            {
                using (var target = new SpreadsheetIndexer(spreadsheetBytes))
                {
                    Assert.IsNotNull(target);
                    Assert.IsNotNull(target.Workbook);
                    Assert.IsTrue(target.Workbook.Worksheets.Count() == 3);
                }
            }
        }

        [TestMethod]
        [DeploymentItem(RandomDataThreeSheetSpath, TestFilesDir)]
        public void Constructor_RandomDataMultiSheets_ValidState()
        {
            var spreadsheetBytes = LoadTestSpreadSheetBytes(RandomDataThreeSheetSpath);
            if (spreadsheetBytes != null)
            {
                using (var target = new SpreadsheetIndexer(spreadsheetBytes))
                {
                    Assert.IsNotNull(target);
                    Assert.IsNotNull(target.Workbook);
                    Assert.IsTrue(target.Workbook.Worksheets.Count() == 3);
                }
            }
        }

        [TestMethod]
        [DeploymentItem(EmptySheetPath, TestFilesDir)]
        public void Constructor_EmptySheetStream_ValidState()
        {
            AssertFileExists(EmptySheetPath);
            using (var target = new SpreadsheetIndexer(File.Open(EmptySheetPath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite)))
            {
                Assert.IsNotNull(target);
                Assert.IsNotNull(target.Workbook);
                Assert.IsTrue(target.Workbook.Worksheets.Count() == 1);
            }
        }

        [TestMethod]
        [DeploymentItem(RandomDataSheetSpath, TestFilesDir)]
        public void Constructor_RandomDataStream_ValidState()
        {
            AssertFileExists(RandomDataSheetSpath);
            using (var target = new SpreadsheetIndexer(File.Open(RandomDataSheetSpath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite)))
            {
                Assert.IsNotNull(target);
                Assert.IsNotNull(target.Workbook);
                Assert.IsTrue(target.Workbook.Worksheets.Count() == 1);
            }
        }

        [TestMethod]
        [DeploymentItem(EmptyThreeSheetsPath, TestFilesDir)]
        public void Constructor_EmptyMultiSheetsStream_ValidState()
        {
            AssertFileExists(EmptyThreeSheetsPath);
            using (var target = new SpreadsheetIndexer(File.Open(EmptyThreeSheetsPath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite)))
            {
                Assert.IsNotNull(target);
                Assert.IsNotNull(target.Workbook);
                Assert.IsTrue(target.Workbook.Worksheets.Count() == 3);
            }
        }

        [TestMethod]
        [DeploymentItem(RandomDataThreeSheetSpath, TestFilesDir)]
        public void Constructor_RandomDataMultiSheetsStream_ValidState()
        {
            AssertFileExists(RandomDataThreeSheetSpath);
            using (var target = new SpreadsheetIndexer(File.Open(RandomDataThreeSheetSpath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite)))
            {
                Assert.IsNotNull(target);
                Assert.IsNotNull(target.Workbook);
                Assert.IsTrue(target.Workbook.Worksheets.Count() == 3);
            }
        }

        #endregion
    }
}
