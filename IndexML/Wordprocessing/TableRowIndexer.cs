namespace IndexML.Wordprocessing
{
    using System;
    using DocumentFormat.OpenXml.Wordprocessing;

    /// <summary>
    /// OpenXml utility class for manipulating a table row.
    /// </summary>
    public class TableRowIndexer
    {
        #region Constructors & Destructors

        public TableRowIndexer(TableRow toIndex)
        {
            if (toIndex == null)
            {
                throw new ArgumentNullException("toIndex");
            }

            this.Row = toIndex;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets the wrapped table row element.
        /// </summary>
        public TableRow Row { get; private set; }

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
