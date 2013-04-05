namespace IndexML.Wordprocessing
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using DocumentFormat.OpenXml.Wordprocessing;

    /// <summary>
    /// OpenXml utility class for manipulating a table row.
    /// </summary>
    public class TableRowIndexer
    {
        #region Fields & Constants

        private readonly IList<TableCellIndexer> cells;

        #endregion

        #region Constructors & Destructors

        public TableRowIndexer(TableRow toIndex)
        {
            if (toIndex == null)
            {
                throw new ArgumentNullException("toIndex");
            }

            var cells = toIndex.Elements<TableCell>();
            if (cells.Count() == 0)
            {
                throw new ArgumentException("Invalid table row! A row must contains at least one cell.");
            }

            this.Row = toIndex;
            this.cells = new List<TableCellIndexer>();

            foreach (var cell in cells)
            {
                this.cells.Add(new TableCellIndexer(cell));
            }
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets the wrapped table row element.
        /// </summary>
        public TableRow Row { get; private set; }

        /// <summary>
        /// Gets an enumeration of all the cells in the row.
        /// </summary>
        public IEnumerable<TableCellIndexer> Cells
        {
            get
            {
                return this.cells;
            }
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Casts the indexer to a TableRow element. Any changes made to the raw element will not
        /// be reflected in the indexer.
        /// </summary>
        /// <param name="indexer">The indexer to cast.</param>
        /// <returns>The element that the indexer wraps, or null if <paramref name="indexer"/> is null.</returns>
        public static implicit operator TableRow(TableRowIndexer indexer)
        {
            return indexer != null ? indexer.Row : null;
        }

        #endregion
    }
}
