﻿namespace IndexML.UnitTests.Spreadsheet
{
    using System;
    using System.IO;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using IndexML.Spreadsheet;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A base test class for testing spreadsheet indexers.
    /// </summary>
    public abstract class SpreadsheetTest : Test
    {
        #region Fields & Constants

        protected const string EmptySheetPath = TestFilesDir + "Empty.xlsx";

        protected const string RandomDataSheetSpath = TestFilesDir + "RandomTestData.xlsx";

        protected const string MaxExtentsSheetPath = TestFilesDir + "MaxExtents.xlsx";

        protected const string ExactlyFiveRowsSheetPath = TestFilesDir + "ExactlyFiveRows.xlsx";

        protected const string FiveEvenRowsSheetPath = TestFilesDir + "FiveEvenRows.xlsx";

        protected const string EmptyThreeSheetsPath = TestFilesDir + "EmptyThreeSheet.xlsx";

        protected const string RandomDataThreeSheetSpath = TestFilesDir + "RandomTestDataThreeSheet.xlsx";

        protected const string ColumnValidationsSheetPath = TestFilesDir + "ColumnValidations.xlsx";

        protected const string RowValidationsSheetPath = TestFilesDir + "RowValidations.xlsx";

        protected const string AllValidationsSheetPath = TestFilesDir + "AllValidations.xlsx";

        protected const string OneValidationA2SheetPath = TestFilesDir + "OneValidationCellA2.xlsx";

        protected const string StaggeredValidationsSheetPath = TestFilesDir + "StaggeredValidations.xlsx";

        #endregion

        #region Protected Methods

        protected static SpreadsheetDocument LoadTestSpreadSheet(string path)
        {
            AssertFileExists(path);

            try
            {
                using (var fileStream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    var memory = new MemoryStream(); // Make the stream expandable by using default ctor
                    CopyStream(fileStream, memory);      // Copy the stream to memory so we can do whatever we want with it
                    return SpreadsheetDocument.Open(memory, true);
                }
            }
            catch (Exception exc)
            {
                Assert.Inconclusive("Test inconclusive. Unable to open the document at path " + path + ". Exception: " + exc.Message);
            }

            return null;
        }

        protected static void ValidateRowSequence(ISheetDataIndexer indexer)
        {
            Row previous = null;
            foreach (var current in indexer.Rows)
            {
                if (previous != null)
                {
                    Assert.IsTrue(current.RowIndex > previous.RowIndex);
                }

                previous = current;
            }
        }

        protected static void SafeExecuteTest<TActionable>(
            string spreadsheetPath,
            Func<SpreadsheetDocument, TActionable> selector,
            Action<TActionable> testToPerform)
        {
            if (testToPerform == null)
            {
                Assert.Inconclusive("No test specified to perform!");
            }

            var spreadsheet = LoadTestSpreadSheet(spreadsheetPath);
            if (spreadsheet != null)
            {
                using (spreadsheet)
                {
                    var items = selector == null ? default(TActionable) : selector(spreadsheet);
                    testToPerform(items);
                }
            }
        }

        protected static void SafeExecuteTest(
            string spreadsheetPath, 
            Action<SpreadsheetDocument> testToPerform)
        {
            if (testToPerform == null)
            {
                Assert.Inconclusive("No test specified to perform!");
            }

            var spreadsheet = LoadTestSpreadSheet(spreadsheetPath);
            if (spreadsheet != null)
            {
                using (spreadsheet)
                {
                    testToPerform(spreadsheet);
                }
            }
        }

        #endregion
    }
}
