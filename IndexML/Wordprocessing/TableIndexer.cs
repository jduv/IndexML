namespace IndexML.Wordprocessing
{
    using System;
    using DocumentFormat.OpenXml.Wordprocessing;

    /// <summary>
    /// OpenXml utilty class for indexing tables.
    /// </summary>
    public class TableIndexer
    {
        #region Constructors & Destructors

        public TableIndexer(Table toIndex)
        {
            if (toIndex == null)
            {
                throw new ArgumentNullException("toIndex");
            }

            this.Table = toIndex;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets the table element the indexer wraps.
        /// </summary>
        public Table Table { get; private set; }

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
