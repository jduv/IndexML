namespace IndexML.Spreadsheet
{
    using System;
    using System.Collections.Generic;
    using DocumentFormat.OpenXml.Spreadsheet;

    /// <summary>
    /// OpenXml utiltiy class for indexing rows.
    /// </summary>
    public class RowIndexer
    {
        #region Fields & Constants

        /// <summary>
        /// The capacity of the indexer.
        /// </summary>
        private static readonly short Capacity = 1024 * 16;

        /// <summary>
        /// A dictionary of cells. This can take up a bit of memory.
        /// </summary>        
        private IDictionary<long, CellIndexer> cells = new Dictionary<long, CellIndexer>();

        /// <summary>
        /// The maximum column index for the indexer. Zero based.
        /// </summary>
        private long maxColumnIndex = 0;

        #endregion

        #region Constructors & Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="RowIndexer"/> class.
        /// </summary>
        /// <param name="row">The row to index.</param>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="row"/> is null.</exception>
        public RowIndexer(Row row)
        {
            if (row == null)
            {
                throw new ArgumentNullException("row");
            }

            long count = 0;
            long columnIndex = 0;
            foreach (var cell in row.Descendants<Cell>())
            {
                var cellIndexer = new CellIndexer(cell);
                columnIndex = cellIndexer.ColumnIndex;
                this.cells[columnIndex] = cellIndexer;
                count++;
            }

            this.maxColumnIndex = columnIndex;
            this.Count = count;
            this.Row = row;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets the object associated with this indexer. Changes made to the cell will not be reflected
        /// to any dependent properties inside the indexer, so use this with care.
        /// </summary>
        public Row Row { get; private set; }

        /// <summary>
        /// Gets the number of columns in this row.
        /// </summary>
        public long Count { get; private set; }

        /// <summary>
        /// Gets or sets the index for this row. This property is virtual for testing purposes.
        /// </summary>
        public virtual uint RowIndex
        {
            get
            {
                return this.Row.RowIndex;
            }

            set
            {
                this.Row.RowIndex = value;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this row has any cells or not.
        /// </summary>
        public bool IsEmpty
        {
            get
            {
                return this.Count <= 0;
            }
        }

        /// <summary>
        /// Gets the maximum column index.
        /// </summary>
        /// <exception cref="InvalidOperationException">Thrown if there are no cells in the row.</exception>
        public long MaxColumnIndex
        {
            get
            {
                if (this.IsEmpty)
                {
                    throw new InvalidOperationException("No columns exist in the indexer!");
                }
                else
                {
                    return this.maxColumnIndex + 1;
                }
            }
        }

        /// <summary>
        /// Gets a list of all cells inside the indexer.
        /// </summary>
        public IEnumerable<CellIndexer> Cells
        {
            get
            {
                return this.cells.Values;
            }
        }

        /// <summary>
        /// Gets the cell at the target column index.
        /// </summary>
        /// <param name="colIndex">The column index, which should be one based.</param>
        /// <returns>The cell at the target index, or null if the column doesn't exist in the row.</returns>
        /// <exception cref="IndexOutOfRangeExcception">Thrown if <paramref name="colIndex"/> is
        /// out of range.</exception>
        public CellIndexer this[long colIndex]
        {
            get
            {
                if (colIndex > Capacity || colIndex < 1)
                {
                    throw new IndexOutOfRangeException("Column index out of range!");
                }

                if (this.cells.ContainsKey(colIndex))
                {
                    return this.cells[colIndex];
                }

                return null;
            }
        }

        /// <summary>
        /// Gets the cell at the target column name.
        /// </summary>
        /// <param name="colName">The cell name whose column to retrieve. This can be with or 
        /// without trailing numbers (e.g. AA and AA11 both result in a column index of 27).</param>
        /// <returns>The cell at the target column name, or null if the column doesn't exist in the row.</returns>
        /// <exception cref="ArgumentException">Thrown if <paramref name="colName"/> is null or
        /// malformed.</exception>
        /// <exception cref="IndexOutOfRangeExcception">Thrown if the translated column index is
        /// out of range.</exception>
        public CellIndexer this[string colName]
        {
            get
            {
                if (string.IsNullOrEmpty(colName))
                {
                    throw new ArgumentException("The string parameter cannot be null or empty!", "columnName");
                }

                long colIdx;
                if (CellReference.TryGetColumnIndex(colName, false, out colIdx))
                {
                    return this[colIdx];
                }
                else
                {
                    throw new InvalidOperationException("Unable to parse column index from string " + colName);
                }
            }
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Casts the indexer into a Row object. Any changes made to the result of this
        /// cast will not be reflected in the indexer, so use this with care.
        /// </summary>
        /// <param name="indexer">The indexer to cast.</param>
        /// <returns>The indexer's wrapped object.</returns>
        public static implicit operator Row(RowIndexer indexer)
        {
            return indexer != null ? indexer.Row : null;
        }

        /// <summary>
        /// Clones the underlying row and returns an indexer for it.
        /// </summary>
        /// <returns>A new <see cref="RowIndexer"/> instance on a copy of this one's
        /// Row object.</returns>
        public RowIndexer Clone()
        {
            return new RowIndexer((Row)this.Row.Clone());
        }

        #endregion
    }
}
