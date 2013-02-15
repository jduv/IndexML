namespace IndexML.Spreadsheet
{
    using System.Collections.Generic;
    using DocumentFormat.OpenXml.Spreadsheet;

    /// <summary>
    /// Defines a basic set of behaviors for SheetDataIndexer implementations.
    /// </summary>
    public interface ISheetDataIndexer
    {
        #region Interface Methods


        /// <summary>
        /// Creates a new <see cref="RowIndexer"/> for the target row and appends it to the
        /// end of this indexer. Null arguments will be ignored.
        /// </summary>
        /// <param name="toAppend">The row to append.</param>
        void AppendRow(Row toAppend);

        /// <summary>
        /// Adds a <see cref="RowIndexer"/> to the end of this indexer. Null arguments will be ignored.
        /// </summary>
        /// <param name="toAppend">The row to append.</param>
        void AppendRow(RowIndexer toAppend);

        /// <summary>
        /// Creates a new <see cref="RowIndexer"/> for the target row and inserts it at the given index. 
        /// Null arguments will be ignored.
        /// </summary>
        /// <param name="toInsert">The row to insert.</param>
        /// <param name="rowIndex">The index to insert the row at.</param>
        /// <param name="shiftRowsDown">Should all rows be shifted down after the insert?</param>
        /// <exception cref="IndexOutOfRangeException">Thrown if <paramref name="rowIndex"/>
        /// is out of range.</exception>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="toInsert"/> is null.
        /// </exception>
        void InsertRow(Row toInsert, long rowIndex, bool shiftRowsDown = false);

        /// <summary>
        /// Inserts a <see cref="RowIndexer"/> at the given index. Null arguments will be ignored.
        /// </summary>
        /// <param name="toInsert">The row to insert.</param>
        /// <param name="rowIndex">The index to insert the row at.</param>
        /// <param name="shiftRowsDown">Should all rows be shifted down after the insert?</param>
        /// <exception cref="IndexOutOfRangeException">Thrown if <paramref name="rowIndex"/>
        /// is out of range.</exception>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="toInsert"/> is null.
        /// </exception>
        void InsertRow(RowIndexer toInsert, long rowIndex, bool shiftRowsDown = false);

        /// <summary>
        /// Deletes the row at the target index.
        /// </summary>
        /// <param name="rowIndex">The row index of the row to delete.</param>
        /// <param name="shiftRowsUp">Should all rows be shifted up after the removal?</param>
        /// <returns>True if deletion was successful, false otherwise.</returns>
        /// <exception cref="IndexOutOfRangeException">Thrown if <paramref name="rowIndex"/> is
        /// out of bounds.</exception>
        bool RemoveRow(long rowIndex, bool shiftRowsUp = false);

        /// <summary>
        /// Clones the row at the target index and returns a <see cref="RowIndexer"/> for it.
        /// </summary>
        /// <param name="rowIndex">The index of the row to clone.</param>
        /// <returns>An indexer for the newly cloned row.</returns>
        RowIndexer CloneRow(int rowIndex);

        #endregion

        #region Interface Properties

        /// <summary>
        /// Gets the object associated with this indexer. Changes made to the object will not be reflected
        /// to any dependent properties inside the indexer, so use this with care.
        /// </summary>
        SheetData SheetData { get; }

        /// <summary>
        /// Gets the number of rows in the indexer.
        /// </summary>
        long Count { get; }

        /// <summary>
        /// Gets a value indicating whether the indexer is empty or not.
        /// </summary>
        bool IsEmpty { get; }

        /// <summary>
        /// Gets the maximum row index.
        /// </summary>
        /// <exception cref="InvalidOperationException">Thrown if the indexer is empty.</exception>
        long MaxRowIndex { get; }

        /// <summary>
        /// Gets an enumeration of all the rows inside the indexer.
        /// </summary>
        IEnumerable<RowIndexer> Rows { get; }

        /// <summary>
        /// Gets the row at the target index. This operation is O(n).
        /// </summary>
        /// <param name="rowIndex">The index of the row to retrieve.</param>
        /// <returns>The row at the given index.</returns>
        /// <exception cref="IndexOutOfRangeException">Thrown if <paramref name="rowIndex"/> is out of bounds.</exception>
        RowIndexer this[long rowIndex] { get; }

        #endregion
    }
}
