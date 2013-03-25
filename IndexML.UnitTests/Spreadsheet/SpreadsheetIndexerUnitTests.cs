namespace IndexML.UnitTests.Spreadsheet
{
    using System;
    using System.IO;
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using IndexML.Spreadsheet;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    public class SpreadsheetIndexerUnitTests : SpreadsheetTest
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
            var spreadsheetBytes = LoadTestFileBytes(EmptySheetPath);
            if (spreadsheetBytes != null)
            {
                using (var target = new SpreadsheetIndexer(spreadsheetBytes))
                {
                    Assert.IsNotNull(target);
                    Assert.IsNotNull(target.Workbook);
                    Assert.IsTrue(target.Workbook.Worksheets.Count() == 1);
                    Assert.IsFalse(target.Disposed);

                    // Check properties
                    Assert.IsNotNull(target.Data);
                    Assert.IsTrue(target.Data.Length > 0);
                    Assert.IsTrue(target.Bytes.Length > 0);
                }
            }
        }

        [TestMethod]
        [DeploymentItem(RandomDataSheetSpath, TestFilesDir)]
        public void Constructor_RandomData_ValidState()
        {
            var spreadsheetBytes = LoadTestFileBytes(RandomDataSheetSpath);
            if (spreadsheetBytes != null)
            {
                using (var target = new SpreadsheetIndexer(spreadsheetBytes))
                {
                    Assert.IsNotNull(target);
                    Assert.IsNotNull(target.Workbook);
                    Assert.IsTrue(target.Workbook.Worksheets.Count() == 1);
                    Assert.AreEqual(spreadsheetBytes.Length, target.Bytes.Length);
                    Assert.IsFalse(target.Disposed);

                    // Check properties
                    Assert.IsNotNull(target.Data);
                    Assert.IsTrue(target.Data.Length > 0);
                    Assert.IsTrue(target.Bytes.Length > 0);
                }
            }
        }

        [TestMethod]
        [DeploymentItem(EmptyThreeSheetsPath, TestFilesDir)]
        public void Constructor_EmptyMultiSheets_ValidState()
        {
            var spreadsheetBytes = LoadTestFileBytes(EmptyThreeSheetsPath);
            if (spreadsheetBytes != null)
            {
                using (var target = new SpreadsheetIndexer(spreadsheetBytes))
                {
                    Assert.IsNotNull(target);
                    Assert.IsNotNull(target.Workbook);
                    Assert.IsTrue(target.Workbook.Worksheets.Count() == 3);
                    Assert.IsFalse(target.Disposed);

                    // Check properties
                    Assert.IsNotNull(target.Data);
                    Assert.IsTrue(target.Data.Length > 0);
                    Assert.IsTrue(target.Bytes.Length > 0);
                }
            }
        }

        [TestMethod]
        [DeploymentItem(RandomDataThreeSheetSpath, TestFilesDir)]
        public void Constructor_RandomDataMultiSheets_ValidState()
        {
            var spreadsheetBytes = LoadTestFileBytes(RandomDataThreeSheetSpath);
            if (spreadsheetBytes != null)
            {
                using (var target = new SpreadsheetIndexer(spreadsheetBytes))
                {
                    Assert.IsNotNull(target);
                    Assert.IsNotNull(target.Workbook);
                    Assert.IsTrue(target.Workbook.Worksheets.Count() == 3);
                    Assert.IsFalse(target.Disposed);

                    // Check properties
                    Assert.IsNotNull(target.Data);
                    Assert.IsTrue(target.Data.Length > 0);
                    Assert.IsTrue(target.Bytes.Length > 0);
                }
            }
        }

        [TestMethod]
        [DeploymentItem(EmptySheetPath, TestFilesDir)]
        public void Constructor_EmptySheetStream_ValidState()
        {
            AssertFileExists(EmptySheetPath);
            using (var target = new SpreadsheetIndexer(
                File.Open(EmptySheetPath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite)))
            {
                Assert.IsNotNull(target);
                Assert.IsNotNull(target.Workbook);
                Assert.IsTrue(target.Workbook.Worksheets.Count() == 1);
                Assert.IsFalse(target.Disposed);

                // Check properties
                Assert.IsNotNull(target.Data);
                Assert.IsTrue(target.Data.Length > 0);
                Assert.IsTrue(target.Bytes.Length > 0);
            }
        }

        [TestMethod]
        [DeploymentItem(RandomDataSheetSpath, TestFilesDir)]
        public void Constructor_RandomDataStream_ValidState()
        {
            AssertFileExists(RandomDataSheetSpath);
            using (var target = new SpreadsheetIndexer(
                File.Open(RandomDataSheetSpath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite)))
            {
                Assert.IsNotNull(target);
                Assert.IsNotNull(target.Workbook);
                Assert.IsTrue(target.Workbook.Worksheets.Count() == 1);
                Assert.IsFalse(target.Disposed);

                // Check properties
                Assert.IsNotNull(target.Data);
                Assert.IsTrue(target.Data.Length > 0);
                Assert.IsTrue(target.Bytes.Length > 0);
            }
        }

        [TestMethod]
        [DeploymentItem(EmptyThreeSheetsPath, TestFilesDir)]
        public void Constructor_EmptyMultiSheetsStream_ValidState()
        {
            AssertFileExists(EmptyThreeSheetsPath);
            using (var target = new SpreadsheetIndexer(
                File.Open(EmptyThreeSheetsPath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite)))
            {
                Assert.IsNotNull(target);
                Assert.IsNotNull(target.Workbook);
                Assert.IsTrue(target.Workbook.Worksheets.Count() == 3);
                Assert.IsFalse(target.Disposed);

                // Check properties
                Assert.IsNotNull(target.Data);
                Assert.IsTrue(target.Data.Length > 0);
                Assert.IsTrue(target.Bytes.Length > 0);
            }
        }

        [TestMethod]
        [DeploymentItem(RandomDataThreeSheetSpath, TestFilesDir)]
        public void Constructor_RandomDataMultiSheetsStream_ValidState()
        {
            AssertFileExists(RandomDataThreeSheetSpath);
            using (var target = new SpreadsheetIndexer(
                File.Open(RandomDataThreeSheetSpath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite)))
            {
                Assert.IsNotNull(target);
                Assert.IsNotNull(target.Workbook);
                Assert.IsTrue(target.Workbook.Worksheets.Count() == 3);
                Assert.IsFalse(target.Disposed);

                // Check properties
                Assert.IsNotNull(target.Data);
                Assert.IsTrue(target.Data.Length > 0);
                Assert.IsTrue(target.Bytes.Length > 0);
            }
        }

        [TestMethod]
        [DeploymentItem(RandomDataThreeSheetSpath, TestFilesDir)]
        public void SaveAndClose_DisposesIndexer()
        {
            AssertFileExists(RandomDataThreeSheetSpath);
            var target = new SpreadsheetIndexer(
                File.Open(RandomDataThreeSheetSpath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite));
            target.SaveAndClose();

            Assert.IsTrue(target.Disposed);
        }

        [TestMethod]
        [DeploymentItem(RandomDataThreeSheetSpath, TestFilesDir)]
        public void SaveAndReopen_DoesNotDisposeIndexer()
        {
            AssertFileExists(RandomDataThreeSheetSpath);
            using (var target = new SpreadsheetIndexer(
                File.Open(RandomDataThreeSheetSpath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite)))
            {
                target.SaveAndReopen();
                Assert.IsFalse(target.Disposed);
            }
        }

        [TestMethod]
        [DeploymentItem(RandomDataThreeSheetSpath, TestFilesDir)]
        [ExpectedException(typeof(ObjectDisposedException))]
        public void DataProperty_DisposedObject()
        {
            AssertFileExists(RandomDataThreeSheetSpath);
            var target = new SpreadsheetIndexer(
                File.Open(RandomDataThreeSheetSpath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite));
            target.SaveAndClose();

            Assert.IsTrue(target.Disposed);
            var data = target.Data;
        }

        [TestMethod]
        [DeploymentItem(RandomDataThreeSheetSpath, TestFilesDir)]
        [ExpectedException(typeof(ObjectDisposedException))]
        public void BytesProperty_DisposedObject()
        {
            AssertFileExists(RandomDataThreeSheetSpath);
            var target = new SpreadsheetIndexer(
                File.Open(RandomDataThreeSheetSpath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite));
            target.SaveAndClose();

            Assert.IsTrue(target.Disposed);
            var data = target.Bytes;
        }

        [TestMethod]
        [DeploymentItem(RandomDataThreeSheetSpath, TestFilesDir)]
        public void ImplicitCast_ValidIndexerSameReference()
        {
            AssertFileExists(RandomDataThreeSheetSpath);
            using (var target = new SpreadsheetIndexer(
                File.Open(RandomDataThreeSheetSpath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite)))
            {
                SpreadsheetDocument spreadsheet = (SpreadsheetDocument)target;
                Assert.IsNotNull(spreadsheet);
                Assert.AreSame(target.SpreadsheetDocument, spreadsheet);
            }
        }

        #endregion
    }
}
