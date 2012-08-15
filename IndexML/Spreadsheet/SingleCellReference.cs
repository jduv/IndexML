﻿namespace IndexML
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
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="cellRef"/> is null.</exception>
        public override bool ContainsOrSubsumes(ICellReference cellRef)
        {
            if (cellRef == null)
            {
                throw new ArgumentNullException("cellRef");
            }

            return this.Value.Equals(cellRef.Value, StringComparison.OrdinalIgnoreCase);
        }

        #endregion
    }
}
