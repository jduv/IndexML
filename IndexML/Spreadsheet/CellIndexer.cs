namespace IndexML
{
    using System;
    using DocumentFormat.OpenXml.Spreadsheet;

    /// <summary>
    /// OpenXml utility class for getting values from a Cell.
    /// </summary>
    public class CellIndexer
    {
        #region Constructors & Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="CellIndexer"/> class.
        /// </summary>
        /// <param name="cell">The cell to initialize with.</param>
        /// <exception cref="ArgumentNullException">Thrown if <see cref="cell"/> is null.</exception>
        public CellIndexer(Cell cell)
        {
            if (cell == null)
            {
                throw new ArgumentNullException("cell");
            }

            this.Cell = cell;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets the object associated with this indexer. Changes made to the cell will not be reflected
        /// to any dependent properties inside the indexer, so use this with care.
        /// </summary>
        public Cell Cell { get; private set; }

        /// <summary>
        /// Gets or setsa type for this cell. This will be null if the cell has no valid 
        /// data type object associated with it.
        /// </summary>
        public CellValues? DataType
        {
            get
            {
                if (this.Cell.DataType != null)
                {
                    return this.Cell.DataType;
                }
                else
                {
                    return null;
                }
            }

            set
            {
                this.Cell.DataType = value;
            }
        }

        /// <summary>
        /// Gets the column name for this cell. This is generated on every call, so cache it
        /// if you're using it often.
        /// </summary>
        public string ColumnName
        {
            get
            {
                return GetColumnName(this.Cell);
            }
        }

        /// <summary>
        /// Gets the column index for this cell. This is generated on every call, so cache it
        /// if you're using it often.
        /// </summary>
        public long ColumnIndex
        {
            get
            {
                return GetColumnIndex(this.Cell);
            }
        }

        /// <summary>
        /// Gets or sets the value for the cell.
        /// </summary>
        public string Value
        {
            get
            {
                if (this.Cell.CellValue != null)
                {
                    return this.Cell.CellValue.Text;
                }

                return null;
            }

            set
            {
                if (value == null)
                {
                    this.Cell.CellValue = null;
                }
                else if (this.Cell.CellValue == null)
                {
                    this.Cell.CellValue = new CellValue() { Text = value };
                }
                else
                {
                    this.Cell.CellValue.Text = value;
                }
            }
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Casts the indexer into a Cell object. Any changes made to the result of this
        /// cast will not be reflected in the indexer, so use this with care.
        /// </summary>
        /// <param name="indexer">The indexer to cast.</param>
        /// <returns>The indexer's wrapped object.</returns>
        public static implicit operator Cell(CellIndexer indexer)
        {
            return indexer != null ? indexer.Cell : null;
        }

        /// <summary>
        /// Parses the column reference for the target cell and returns it as an index.
        /// </summary>
        /// <param name="cell">The cell whose index to retrieve.</param>
        /// <returns>The column index for the target cell.</returns>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="cell"/> is null.</exception>
        /// <exception cref="ArgumentException">Thrown if <paramref name="cell"/> has a missing or 
        /// invalid cell reference.</exception>
        public static long GetColumnIndex(Cell cell)
        {
            if (cell == null)
            {
                throw new ArgumentNullException("cell");
            }

            if (cell.CellReference == null)
            {
                throw new ArgumentException("Invalid cell reference on the target cell, unable to parse column index.");
            }

            long colIdx;
            if (CellReference.TryGetColumnIndex(cell.CellReference.Value, true, out colIdx))
            {
                return colIdx;
            }
            else
            {
                throw new ArgumentException("Unable to parse column index for the cell reference " + cell.CellReference.Value);
            }
        }

        /// <summary>
        /// Gets the column name for the target cell, stripping all row references from it.
        /// </summary>
        /// <param name="cell">The cell whose column name to retrieve.</param>
        /// <returns>The column name of the target cell.</returns>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="cell"/> is null.</exception>
        /// <exception cref="ArgumentException">Thrown if <paramref name="cell"/>'s cell reference is 
        /// empty or malformed.</exception>
        public static string GetColumnName(Cell cell)
        {
            if (cell == null)
            {
                throw new ArgumentNullException("cell");
            }

            if (cell.CellReference == null)
            {
                throw new ArgumentException("Invalid cell reference on the target cell, unable to parse column name.");
            }

            string columnName;
            if (CellReference.TryGetColumnName(cell.CellReference.Value, true, out columnName))
            {
                return columnName;
            }

            throw new ArgumentException("Malformed cell reference on the target cell, unable to parse column name.");
        }

        #endregion
    }
}
