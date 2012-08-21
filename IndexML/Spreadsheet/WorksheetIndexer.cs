﻿namespace IndexML
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;

    /// <summary>
    /// OpenXml utility class for managing a worksheet.
    /// </summary>
    public class WorksheetIndexer
    {
        #region Fields & Constants

        /// <summary>
        /// The sheet associated with the worksheet.
        /// </summary>
        private readonly Sheet sheet;

        /// <summary>
        /// A list of data validations.
        /// </summary>
        private readonly IList<DataValidation> dataValidations = new List<DataValidation>();

        #endregion

        #region Constructors & Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="WorksheetIndexer"/> class.
        /// </summary>
        /// <param name="toIndex">The worksheet part to initialize with.</param>
        /// <param name="sheet">The sheet to associate with the worksheet. Sheet objects contain some extra metadata
        /// about the worksheet, so if passed the indexer will reflect this extra data.</param>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="toIndex"/> is null.</exception>
        public WorksheetIndexer(WorksheetPart toIndex, Sheet sheet)
        {
            if (toIndex == null)
            {
                throw new ArgumentNullException("toIndex");
            }

            // Set the sheet. 
            this.sheet = sheet;

            // Index the sheet data.
            var sheetData = toIndex.Worksheet.Descendants<SheetData>().FirstOrDefault();
            this.SheetData = new SheetDataIndexer(sheetData);

            // Add the validators.
            this.dataValidations = toIndex.Worksheet.Descendants<DataValidation>().ToList();

            // Set the worksheet.
            this.Worksheet = toIndex.Worksheet;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets a <see cref="SheetDataIndexer"/> for the worksheet.
        /// </summary>
        public SheetDataIndexer SheetData { get; private set; }

        /// <summary>
        /// Gets the worksheet part that's being indexed.
        /// </summary>
        public Worksheet Worksheet { get; private set; }

        /// <summary>
        /// Gets a list of data validation elements.
        /// </summary>
        public IEnumerable<DataValidation> DataValidations
        {
            get
            {
                return this.dataValidations;
            }
        }

        /// <summary>
        /// Gets the sheet's name.
        /// </summary>
        public string SheetName
        {
            get
            {
                return this.sheet != null ? this.sheet.Name : null;
            }
        }

        #endregion
    }
}
