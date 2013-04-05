namespace IndexML.Wordprocessing
{
    using System;
    using System.Linq;
    using System.Collections.Generic;
    using DocumentFormat.OpenXml.Wordprocessing;

    /// <summary>
    /// OpenXml utilty class for indexing tables.
    /// </summary>
    public class TableIndexer
    {
        #region Fields & Constants

        private IList<TableRowIndexer> tableRows;
        private TableGrid tableGrid;

        #endregion

        #region Constructors & Destructors

        public TableIndexer(Table toIndex)
        {
            if (toIndex == null)
            {
                throw new ArgumentNullException("toIndex");
            }

            this.tableGrid = toIndex.Elements<TableGrid>().FirstOrDefault();
            if (this.tableGrid == null)
            {
                throw new ArgumentException("Invalid table structure! Unable to locate the table grid element.");
            }

            this.Table = toIndex;

            // Fill out the rows.
            this.tableRows = new List<TableRowIndexer>();
            foreach(var row in toIndex.Elements<TableRow>())
            {
                this.tableRows.Add(new TableRowIndexer(row));
            }
            
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets the table element the indexer wraps.
        /// </summary>
        public Table Table { get; private set; }

        /// <summary>
        /// Gets the columns.
        /// </summary>
        public IEnumerable<GridColumn> Columns
        {
            get
            {
                return this.tableGrid.Elements<GridColumn>();
            }
        }

        /// <summary>
        /// Gets the rows.
        /// </summary>
        public IEnumerable<TableRowIndexer> Rows
        {
            get 
            {
                return this.tableRows;
            }
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Casts the indexer to a Table element. Any changes made to the raw element will not
        /// be reflected in the indexer.
        /// </summary>
        /// <param name="indexer">The indexer to cast.</param>
        /// <returns>The element that the indexer wraps, or null if <paramref name="indexer"/> is null.</returns>
        public static implicit operator Table(TableIndexer indexer)
        {
            return indexer != null ? indexer.Table : null;
        }

        #endregion
    }
}
