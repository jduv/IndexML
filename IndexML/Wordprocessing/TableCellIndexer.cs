namespace IndexML.Wordprocessing
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using DocumentFormat.OpenXml.Wordprocessing;

    /// <summary>
    /// OpenXml utility class for manipulating table cells.
    /// </summary>
    public class TableCellIndexer
    {
        #region Constructors & Destructors

        public TableCellIndexer(TableCell toIndex)
        {
            if (toIndex == null)
            {
                throw new ArgumentNullException("toIndex");
            }

            this.Cell = toIndex;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets the wrapped cell.
        /// </summary>
        public TableCell Cell { get; private set; }

        #endregion

        #region Public Methods

        /// <summary>
        /// Casts the indexer to a TableCell element. Any changes made to the raw element will not
        /// be reflected in the indexer.
        /// </summary>
        /// <param name="indexer">The indexer to cast.</param>
        /// <returns>The element that the indexer wraps, or null if <paramref name="indexer"/> is null.
        /// </returns>
        public static implicit operator TableCell(TableCellIndexer indexer)
        {
            return indexer != null ? indexer.Cell : null;
        }

        #endregion
    }
}
