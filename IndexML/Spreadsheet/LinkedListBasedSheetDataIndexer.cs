namespace IndexML.Spreadsheet
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text.RegularExpressions;
    using DocumentFormat.OpenXml.Spreadsheet;

    public class LinkedListBasedSheetDataIndexer : ISheetDataIndexer
    {
        #region Fields & Constants

        /// <summary>
        /// Maximum capacity of a sheet.
        /// </summary>
        public static readonly long Capacity = 1024 * 1024;

        /// <summary>
        /// A linked list of row indexers. This will make insertion/deletion/shifting a breeze,
        /// but random access will be O(n).
        /// </summary>
        private LinkedList<RowIndexer> rows = new LinkedList<RowIndexer>();

        /// <summary>
        /// The maximum row index for the indexer. Zero based.
        /// </summary>
        private long maxRowIndex = 0;

        #endregion

        #region Constructors & Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ArrayBasedSheetDataIndexer"/> class.
        /// </summary>
        /// <param name="sheetData">The sheet data to index.</param>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="sheetData"/> is null.</exception>
        public LinkedListBasedSheetDataIndexer(SheetData sheetData)
        {
            if (sheetData == null)
            {
                throw new ArgumentNullException("sheetData");
            }

            long rowIndex = 0;

            // Order this sucker first so we don't have to sort a Linked List.
            foreach (var row in sheetData.Descendants<Row>().OrderBy(x => (uint)x.RowIndex))
            {
                var rowIndexer = new RowIndexer(row);
                rowIndex = rowIndexer.RowIndex;
                this.rows.AddLast(rowIndexer);
            }

            this.maxRowIndex = rowIndex;
            this.SheetData = sheetData;
        }

        #endregion

        #region Properties

        /// <inheritdoc />
        public SheetData SheetData { get; private set; }

        /// <inheritdoc />
        public long Count
        {
            get
            {
                return this.rows.Count;
            }
        }

        /// <inheritdoc />
        public bool IsEmpty
        {
            get
            {
                return this.Count <= 0;
            }
        }

        /// <inheritdoc />
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
                    return this.maxRowIndex;
                }
            }
        }

        /// <inheritdoc />
        public IEnumerable<RowIndexer> Rows
        {
            get
            {
                return this.rows;
            }
        }

        /// <inheritdoc />
        public RowIndexer this[long rowIndex]
        {
            get
            {
                if (rowIndex > Capacity || rowIndex < 1)
                {
                    throw new IndexOutOfRangeException("Row index out of range!");
                }

                return this.rows.FirstOrDefault(x => x.RowIndex == rowIndex);
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
        public static implicit operator SheetData(LinkedListBasedSheetDataIndexer indexer)
        {
            return indexer != null ? indexer.SheetData : null;
        }

        /// <inheritdoc />
        public void AppendRow(Row toAppend)
        {
            if (toAppend != null)
            {
                this.AppendRow(new RowIndexer(toAppend));
            }
        }

        /// <inheritdoc />
        public void AppendRow(RowIndexer toAppend)
        {
            if (toAppend != null)
            {
                this.InsertRow(toAppend, this.maxRowIndex + 1);
            }
        }

        /// <inheritdoc />
        public void InsertRow(Row toInsert, long rowIndex, bool shiftRowsDown = false)
        {
            if (toInsert == null)
            {
                throw new ArgumentNullException("toInsert");
            }

            this.InsertRow(new RowIndexer(toInsert), rowIndex, shiftRowsDown);
        }

        /// <inheritdoc />
        public void InsertRow(RowIndexer toInsert, long rowIndex, bool shiftRowsDown = false)
        {
            // Check for out of bounds and over capacity
            if (rowIndex < 0 || rowIndex > Capacity || this.Count + 1 > Capacity)
            {
                throw new IndexOutOfRangeException("The requested row index is out of bounds!");
            }

            if (toInsert == null)
            {
                throw new ArgumentNullException("toInsert");
            }

            // First, some housekeeping
            toInsert.RowIndex = (uint)rowIndex;
            SyncCellReferencesToRowIndex(toInsert);

            // Easy case. Just add it to the end.
            if (rowIndex > this.maxRowIndex)
            {
                this.rows.AddLast(toInsert);
                this.SheetData.AppendChild<Row>(toInsert);
                this.maxRowIndex = rowIndex;
            }
            else if (shiftRowsDown)
            {
                // Handle shift
                this.InsertAndShiftRowsDown(toInsert, rowIndex);
            }
            else
            {
                LinkedListNode<RowIndexer> afterOrAtIndex = this.FindAfterBeforeOrAt(rowIndex);
                if (afterOrAtIndex.Value.RowIndex == rowIndex)
                {
                    // caught the node at the index, replace
                    this.SheetData.ReplaceChild<Row>(toInsert, afterOrAtIndex.Value);
                    afterOrAtIndex.Value = toInsert;
                }
                else
                {
                    // caught the first node after the index, insert
                    this.SheetData.InsertBefore<Row>(toInsert, afterOrAtIndex.Value);
                    this.rows.AddBefore(afterOrAtIndex, toInsert);
                }
            }
        }

        /// <inheritdoc />
        public bool RemoveRow(long rowIndex, bool shiftRowsUp = false)
        {
            // Check for out of bounds
            if (rowIndex < 0 || rowIndex > Capacity)
            {
                throw new IndexOutOfRangeException("The requested row index is out of bounds!");
            }

            throw new NotImplementedException();
        }

        /// <inheritdoc />
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

        /// <summary>
        /// This method is an optimization over searching in the linked list for
        /// a target row. Why didn't I use linq? Because there's a special terminating
        /// condition that can prevent us from spinning through the entire list, and I want
        /// to return the Node, not the value. 
        /// </summary>
        /// <param name="rowIndex">The index to retrieve.</param>
        /// <returns>The row at the target index, if it exists, or the node immediately
        /// before it in the index sequence.</returns>
        private LinkedListNode<RowIndexer> FindAfterBeforeOrAt(long rowIndex)
        {
            LinkedListNode<RowIndexer> target = null;
            for (target = this.rows.First; target != null; target = target.Next)
            {
                if (target.Value.RowIndex >= rowIndex)
                {
                    // break and return target.
                    break;
                }
            }

            return target;
        }

        /// <summary>
        /// Inserts a row at the target index and shifts all the previous indices down by one.
        /// </summary>
        /// <param name="toInsert">The row to insert.</param>
        /// <param name="rowIndex">The index in which to insert it.</param>
        private void InsertAndShiftRowsDown(RowIndexer toInsert, long rowIndex)
        {
            // Start at the end, work our way backwards
            var current = this.rows.Last;
            while (current.Value.RowIndex >= rowIndex && current != this.rows.First)
            {
                current.Value.RowIndex++;
                SyncCellReferencesToRowIndex(current.Value);
                current = current.Previous;
            }

            // Inserting at the first row is a special case
            if (current == this.rows.First && rowIndex == 1)
            {
                this.rows.First.Value.RowIndex++;
                this.rows.AddFirst(toInsert);
                this.SheetData.InsertBefore<Row>(toInsert, current.Value);
            }
            else
            {
                this.rows.AddAfter(current, toInsert);
                this.SheetData.InsertAfter<Row>(toInsert, current.Value);
            }

            this.maxRowIndex++;
        }

        #endregion
    }
}
