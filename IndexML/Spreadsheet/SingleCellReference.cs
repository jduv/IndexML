namespace IndexML
{
    using System;

    /// <summary>
    /// Represents a single cell reference.
    /// </summary>
    public class SingleCellReference : CellReference
    {
        #region Fields & Constants

        /// <summary>
        /// The column name for the cell reference.
        /// </summary>
        private readonly string colName;

        /// <summary>
        /// The column index for the cell reference.
        /// </summary>
        private readonly long colIdx;

        /// <summary>
        /// The row index for the cell reference.
        /// </summary>
        private readonly long rowIdx;

        #endregion

        #region Constructors & Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="SingleCellReference"/> class.
        /// </summary>
        /// <param name="cellRef">The cell reference to create the reference for.</param>
        public SingleCellReference(string cellRef)
            : base(cellRef)
        {
            // base will throw if cellRef is bad--assume it's good to go.
            if (!TryGetColumnName(cellRef, true, out colName) ||
                !TryGetColumnIndex(cellRef, true, out colIdx) ||
                !TryGetRowIndex(cellRef, out rowIdx))                
            {
                throw new ArgumentException("Unable to parse the cell reference.");
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SingleCellReference"/> class.
        /// </summary>
        /// <param name="row">The row index to initialize with.</param>
        /// <param name="column">The column index to initialize with.</param>
        public SingleCellReference(long row, long column)
        {
            if (column <= 0)
            {
                throw new ArgumentOutOfRangeException("column");
            }

            if (row <= 0)
            {
                throw new ArgumentOutOfRangeException("row");
            }

            this.colIdx = column;
            this.rowIdx = row;
            this.colName = GetColumnName(colIdx);
            this.Value = this.ColumnName + this.RowIndex;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets the column name for the cell reference.
        /// </summary>
        public string ColumnName 
        {
            get
            {
                return colName;
            }
        }

        /// <summary>
        /// Gets the column for the cell reference.
        /// </summary>
        public long ColumnIndex 
        { 
            get
            {
                return colIdx;
            }
        }

        /// <summary>
        /// Gets the row for the cell reference.
        /// </summary>
        public long RowIndex 
        { 
            get
            {
                return rowIdx;
            }
        }

        #endregion

        #region Public Methods
        
        /// <inheritdoc />        
        public override bool ContainsOrSubsumes(ICellReference cellRef)
        {
            if (cellRef == null)
            {
                throw new ArgumentNullException("cellRef");
            }

            return this.Value.Equals(cellRef.Value, StringComparison.OrdinalIgnoreCase);
        }

        /// <inheritdoc/>
        /// <remarks>
        /// Scaling a single cell reference will always result in a ranged cell reference. For example,
        /// if you scale the single reference A1 by (2, 0) the result will be A1:A3. This works in both
        /// directions.
        /// </remarks>
        public override ICellReference Scale(int row, int col)
        {
            var rowIdx = Math.Max(1, this.RowIndex + row);
            var colIdx = Math.Max(1, this.ColumnIndex + col);

            // If no changes were made, i.e. we're at origin scaling negative or (0, 0) was passed
            if (rowIdx == this.RowIndex && colIdx == this.ColumnIndex)
            {
                return new SingleCellReference(this.RowIndex, this.ColumnIndex);
            }
            else
            {
                var start = new SingleCellReference(
                    Math.Min(this.RowIndex, rowIdx), 
                    Math.Min(this.ColumnIndex, colIdx));

                var end = new SingleCellReference(
                    Math.Max(this.RowIndex, rowIdx),
                    Math.Max(this.ColumnIndex, colIdx));

                return new RangeCellReference(start, end);
            }
        }

        /// <inheritdoc />
        /// <remarks>
        /// Translating a cell reference simply moves it as you might expect. The method will not
        /// affect the type of cell reference. For example, translating the cell A1 by (2, 2) will result
        /// in C3.
        /// </remarks>
        public override ICellReference Translate(int row, int col)
        {
            var rowIdx = Math.Max(1, this.RowIndex + row);
            var colIdx = Math.Max(1, this.ColumnIndex + col);

            return new SingleCellReference(rowIdx, colIdx);
        }      

        #endregion
    }
}
