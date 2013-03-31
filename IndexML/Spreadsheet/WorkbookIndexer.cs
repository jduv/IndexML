namespace IndexML.Spreadsheet
{
    using System;
    using System.Collections.Generic;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;

    /// <summary>
    /// OpenXml utility class for managing a workbook.
    /// </summary>
    public class WorkbookIndexer
    {
        #region Fields & Constants

        /// <summary>
        /// A list of worksheet indexers.
        /// </summary>
        private IList<WorksheetIndexer> worksheets = new List<WorksheetIndexer>();

        #endregion

        #region Constructors & Destructors

        public WorkbookIndexer(WorkbookPart toIndex)
        {
            if (toIndex == null)
            {
                throw new ArgumentNullException("workbookPart");
            }

            foreach (Sheet sheet in toIndex.Workbook.Sheets)
            {
                var worksheet = (WorksheetPart)toIndex.GetPartById(sheet.Id);
                if (worksheet != null)
                {
                    this.worksheets.Add(new WorksheetIndexer(worksheet, sheet));
                }
            }

            // TODO: This should be modified to create the shared string table part instead of just bailing on an empty worksheet.
            this.SharedStringTable = toIndex.SharedStringTablePart != null ? 
                new SharedStringTableIndexer(toIndex.SharedStringTablePart) :
                null;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets the shared string table for the workbook.
        /// </summary>
        public SharedStringTableIndexer SharedStringTable { get; private set; }

        /// <summary>
        /// Gets a list of <see cref="WorksheetIndexer"/> objects for the workbook.
        /// </summary>
        public IEnumerable<WorksheetIndexer> Worksheets
        {
            get
            {
                return this.worksheets;
            }
        }

        #endregion
    }
}
