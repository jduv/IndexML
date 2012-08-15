namespace IndexML
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text.RegularExpressions;
    using DocumentFormat.OpenXml.Spreadsheet;
    using IndexML.Extensions;

    /// <summary>
    /// OpenXml utility class for manipulating sheet data.
    /// </summary>
    public class SheetDataIndexer
    {
        #region Fields & Constants

        /// <summary>
        /// Maximum capacity of a sheet.
        /// </summary>
        public static readonly long Capacity = 1024 * 1024;

        /// <summary>
        /// An array of rows.
        /// </summary>
        private RowIndexer[] rows = new RowIndexer[Capacity];

        /// <summary>
        /// The maximum row index, zero based.
        /// </summary>
        private long maxRowIndex;

        #endregion

        #region Constructors & Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="SheetDataIndexer"/> class.
        /// </summary>
        /// <param name="sheetData">The sheet data to index.</param>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="sheetData"/> is null.</exception>
        public SheetDataIndexer(SheetData sheetData)
        {
            if (sheetData == null)
            {
                throw new ArgumentNullException("sheetData");
            }

            var rowsToAdd = sheetData.Descendants<Row>().OrderBy(x => (long)x.RowIndex).ToArray();

            if (rowsToAdd.Length > 0)
            {
                foreach (var row in rowsToAdd)
                {
                    this.rows[row.RowIndex - 1] = new RowIndexer(row);
                }

                this.Count = rowsToAdd.Length;
                this.maxRowIndex = rowsToAdd[rowsToAdd.LongLength - 1].RowIndex - 1;
            }
            else
            {
                this.Count = 0;
                this.maxRowIndex = -1;
            }

            this.SheetData = sheetData;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets the object associated with this indexer. Changes made to the cell will not be reflected
        /// to any dependent properties inside the indexer, so use this with care.
        /// </summary>
        public SheetData SheetData { get; private set; }

        /// <summary>
        /// Gets the number of rows in the indexer.
        /// </summary>
        public long Count { get; private set; }

        /// <summary>
        /// Gets a value indicating whether the indexer is empty or not.
        /// </summary>
        public bool IsEmpty
        {
            get
            {
                return this.Count <= 0;
            }
        }

        /// <summary>
        /// Gets the maximum row index. This is based on the actual row index in the sheet data,
        /// not the index into the internal representation. This is one-based.
        /// </summary>
        /// <exception cref="InvalidOperationException">Thrown if the indexer is empty.</exception>
        public long MaxRowIndex
        {
            get
            {
                if (this.IsEmpty)
                {
                    throw new InvalidOperationException("No rows exist in the indexer!");
                }
                else
                {
                    return this.maxRowIndex + 1;
                }
            }
        }

        /// <summary>
        /// Gets an enumeration of all the rows inside the indexer. Returns only valid rows,
        /// no nulls.
        /// </summary>
        public IEnumerable<RowIndexer> Rows
        {
            get
            {
                for (int i = 0; i <= this.maxRowIndex; i++)
                {
                    if (this.rows[i] != null)
                    {
                        yield return this.rows[i];
                    }
                }
            }
        }

        /// <summary>
        /// Gets the row at the target index.
        /// </summary>
        /// <param name="rowIndex">The index of the row to retrieve. This indexer assumes that you are passing
        /// it the row index based on the row object--which is one based--not the internal array.</param>
        /// <returns>The row at the given index.</returns>
        /// <exception cref="IndexOutOfRangeException">Thrown if <paramref name="rowIndex"/> is out of bounds.</exception>
        public RowIndexer this[long rowIndex]
        {
            get
            {
                return this.rows[rowIndex - 1];
            }
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Casts the indexer into a SheetData object. Any changes made to the result of this
        /// cast will not be reflected in the indexer, so use this with care.
        /// </summary>
        /// <param name="indexer">The indexer to convert.</param>
        /// <returns>The indexer's wrapped object.</returns>
        public static implicit operator SheetData(SheetDataIndexer indexer)
        {
            return indexer != null ? indexer.SheetData : null;
        }

        /// <summary>
        /// Creates a new <see cref="RowIndexer"/> for the target row and appends it to the
        /// end of this indexer. Null arguments will be ignored.
        /// </summary>
        /// <param name="toAppend">The row to append.</param>
        public void AppendRow(Row toAppend)
        {
            if (toAppend != null)
            {
                this.InsertRow(new RowIndexer(toAppend), this.maxRowIndex + 2);
            }
        }

        /// <summary>
        /// Adds a <see cref="RowIndexer"/> to the end of this indexer. Null arguments will be ignored.
        /// </summary>
        /// <param name="toAppend">The row to append.</param>
        public void AppendRow(RowIndexer toAppend)
        {
            if (toAppend != null)
            {
                // Add two because row indices are one-based
                this.InsertRow(toAppend, this.maxRowIndex + 2);
            }
        }

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
        public void InsertRow(Row toInsert, long rowIndex, bool shiftRowsDown = false)
        {
            // Check for out of bounds and over capacity
            if (rowIndex < 0 || rowIndex - 1 > Capacity || this.Count + 1 > Capacity)
            {
                throw new IndexOutOfRangeException("The requested row index is out of bounds!");
            }

            if (toInsert == null)
            {
                throw new ArgumentNullException("toInsert");
            }

            this.InsertRow(new RowIndexer(toInsert), rowIndex, shiftRowsDown);
        }

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
        public void InsertRow(RowIndexer toInsert, long rowIndex, bool shiftRowsDown = false)
        {
            // Check for out of bounds and over capacity
            if (rowIndex < 0 || rowIndex - 1 > Capacity || this.Count + 1 > Capacity)
            {
                throw new IndexOutOfRangeException("The requested row index is out of bounds!");
            }

            if (toInsert == null)
            {
                throw new ArgumentNullException("toInsert");
            }

            // First, some housekeeping.
            toInsert.RowIndex = (uint)rowIndex;
            SyncCellReferencesToRowIndex(toInsert);
            long index = rowIndex - 1;
            this.Count = this.rows[index] == null || shiftRowsDown ? this.Count + 1 : this.Count;

            // This catches the empty sheet data case since maxRowIndex is initialized to -1.
            if (index > this.maxRowIndex)
            {
                this.rows[index] = toInsert;
                this.SheetData.AppendChild<Row>(toInsert);
                this.maxRowIndex = index;
            }
            else if (shiftRowsDown)
            {
                // Make room for a new row.
                Row insertBefore = null;
                for (long i = this.maxRowIndex; i >= index; i--)
                {
                    if (this.rows[i] != null)
                    {
                        insertBefore = this.rows[i];
                        insertBefore.RowIndex++;
                        SyncCellReferencesToRowIndex(insertBefore);
                        this.rows.Swap(i, i + 1);
                    }
                }

                this.rows[index] = toInsert;
                this.SheetData.InsertBefore<Row>(toInsert, insertBefore);
                this.maxRowIndex++;
            }
            else
            {
                // If the index we're inserting into isn't null, replace the target child
                if (this.rows[index] != null)
                {
                    this.SheetData.ReplaceChild<Row>(toInsert, this.rows[index]);
                }
                else
                {
                    // Else insert before the node prior to this one
                    Row insertBefore = null;
                    long i = index + 1;

                    do
                    {
                        insertBefore = this.rows[i];
                    }
                    while (insertBefore == null && i <= this.maxRowIndex);

                    if (insertBefore == null)
                    {
                        this.SheetData.AppendChild<Row>(toInsert);
                    }
                    else
                    {
                        this.SheetData.InsertBefore<Row>(toInsert, insertBefore);
                    }
                }

                this.rows[index] = toInsert;
            }
        }

        /// <summary>
        /// Deletes the row at the target index.
        /// </summary>
        /// <param name="rowIndex">The row index of the row to delete.</param>
        /// <param name="shiftRowsUp">Should all rows be shifted up after the removal?</param>
        /// <returns>True if deletion was successful, false otherwise.</returns>
        /// <exception cref="IndexOutOfRangeException">Thrown if <paramref name="rowIndex"/> is
        /// out of bounds.</exception>
        public bool RemoveRow(long rowIndex, bool shiftRowsUp = false)
        {
            // Check for out of bounds
            if (rowIndex < 0 || rowIndex - 1 > Capacity)
            {
                throw new IndexOutOfRangeException("The requested row index is out of bounds!");
            }

            var index = rowIndex - 1;
            if (!this.IsEmpty && this.rows[index] != null)
            {
                var toRemove = this.rows[index];
                this.rows[index] = null;
                this.Count--;

                // If we're not empy, do some book keeping.
                if (!this.IsEmpty)
                {
                    // Are we deleting the last row?
                    if (index == this.maxRowIndex)
                    {
                        // If so, update the max row index.
                        this.maxRowIndex = this.Rows.Last().RowIndex - 1;
                    }
                    else if (shiftRowsUp)
                    {
                        // Else shift all other rows up one
                        for (long i = index + 1; i <= this.maxRowIndex; i++)
                        {
                            var row = this.rows[i];
                            if (row != null)
                            {
                                row.RowIndex--;
                                SyncCellReferencesToRowIndex(row);
                                this.rows.Swap(i, i - 1);
                            }
                        }

                        this.maxRowIndex--;
                    }
                }

                this.SheetData.RemoveChild<Row>(toRemove);
                return true;
            }

            return false;
        }

        /// <summary>
        /// Clones the row at the target index and returns a <see cref="RowIndexer"/> for it.
        /// </summary>
        /// <param name="rowIndex">The index of the row to clone.</param>
        /// <returns>An indexer for the newly cloned row.</returns>
        public RowIndexer CloneRow(int rowIndex)
        {
            var indexerToCopy = this[rowIndex];
            if (indexerToCopy != null)
            {
                var clonedRow = (Row)indexerToCopy.Row.Clone();
                return new RowIndexer(clonedRow);
            }

            return null;
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Synchronizes cell references to a row's index.
        /// </summary>
        /// <param name="toSync">The row to synchronize.</param>        
        private static void SyncCellReferencesToRowIndex(Row toSync)
        {
            foreach (var cell in toSync.Descendants<Cell>())
            {
                cell.CellReference.InnerText = Regex.Replace(
                    cell.CellReference.InnerText,
                    CellReference.SingleCellRefRegexStringStrict,
                    @"${col}" + toSync.RowIndex,
                    RegexOptions.Compiled | RegexOptions.IgnoreCase);
            }

            ValidateCellReferences(toSync);
        }

        /// <summary>
        /// Validates that the cell references in the target row are correctly synchronized to
        /// the index.
        /// </summary>
        /// <param name="toValidate">The row to validate.</param>
        /// <exception cref="ArgumentException">Thrown if the row has some invalid cell references
        /// inside it. This can be fixed by calling <see cref="SyncCellReferencesToRowIndex"/> on the
        /// row.</exception>
        private static void ValidateCellReferences(Row toValidate)
        {
            var cellDict = new Dictionary<string, Cell>();
            foreach (var cell in toValidate.Descendants<Cell>())
            {
                if (cellDict.ContainsKey(cell.CellReference.InnerText))
                {
                    throw new ArgumentException("Duplicate cell reference detected at cell: " + cell.CellReference.InnerText);
                }

                cellDict[cell.CellReference.InnerText] = cell;
            }

            // Make sure the cell row indices are correct.
            var pattern = @"[a-z]+" + toValidate.RowIndex + "{1}";
            var regex = new Regex(pattern, RegexOptions.Compiled | RegexOptions.IgnoreCase);
            foreach (var key in cellDict.Keys)
            {
                if (!regex.IsMatch(key))
                {
                    throw new ArgumentException("Invalid cell reference detected at cell: " + key + ". Regex: " + pattern);
                }
            }
        }

        #endregion
    }
}
