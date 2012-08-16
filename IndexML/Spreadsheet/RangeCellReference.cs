namespace IndexML
{
    using System;

    /// <summary>
    /// Represents a cell reference for a range of cells.
    /// </summary>
    public class RangeCellReference : CellReference
    {
        #region Constructors & Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="RangeCellReference"/> class.
        /// </summary>
        /// <param name="cellRef">The cell to create the reference for.</param>
        public RangeCellReference(string cellRef)
            : base(cellRef)
        {
            // base will throw if cellRef is bad--assume it's good to go.
            var match = RangeCellRefRegex.Match(cellRef);
            if (match.Success)
            {
                this.StartingCellReference = new SingleCellReference(match.Groups["s"].Value);
                this.EndingCellReference = new SingleCellReference(match.Groups["e"].Value);
            }
            else
            {
                // This is probably dead code, base handles this validation--but I'm a pessimist.
                throw new ArgumentException("Invalid string format! It is not a range: " + cellRef);
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RangeCellReference"/> class.
        /// </summary>
        /// <param name="startingCell">The starging cell.</param>
        /// <param name="endingCell">The ending cell.</param>
        /// <exception cref="ArgumentNullException">Thrown if either <paramref name="startingCell"/> or 
        /// <paramref name="endingCell"/> is null.</exception>
        public RangeCellReference(SingleCellReference startingCell, SingleCellReference endingCell)            
        {
            if (startingCell == null)
            {
                throw new ArgumentNullException("startingCell");
            }

            if (endingCell == null)
            {
                throw new ArgumentNullException("endingCell");
            }

            this.StartingCellReference = startingCell;
            this.EndingCellReference = endingCell;
            this.Value = startingCell.Value + ":" + endingCell.Value;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets the starting cell reference for the range.
        /// </summary>
        public SingleCellReference StartingCellReference { get; private set; }

        /// <summary>
        /// Gets the ending cell reference for the range.
        /// </summary>
        public SingleCellReference EndingCellReference { get; private set; }

        /// <summary>
        /// Gets the number of rows contained within the range.
        /// </summary>
        public long Rows
        {
            get
            {
                return this.EndingCellReference.RowIndex;
            }
        }

        /// <summary>
        /// Gets the number of columns contained within the range.
        /// </summary>
        public long Columns
        {
            get
            {
                return this.EndingCellReference.ColumnIndex;
            }
        }

        #endregion

        #region Public Methods

        /// <inheritdoc />
        public override bool ContainsOrSubsumes(ICellReference cellRef)
        {
            if (cellRef is SingleCellReference)
            {
                return this.Contains((SingleCellReference)cellRef);
            }
            else if (cellRef is RangeCellReference)
            {
                return this.Subsumes((RangeCellReference)cellRef);
            }
            else
            {
                return false;
            }
        }

        /// <inheritdoc />
        /// <remarks>
        /// Scaling a ranged cell should make the range larger or smaller based on the arguments
        /// passed. If the range collapses upon itself, this method will return a single cell reference of
        /// the correct resultant cell. For example, if you scale the range A1:C4 by (-2,-2) you should 
        /// get a resulting single cell reference of size 1 pointing to cell A1.
        /// </remarks>
        public override ICellReference Scale(int rows, int cols)
        {
            throw new NotImplementedException();
        }

        /// <inheritdoc />
        /// <remarks>
        /// Translating a ranged cell will move the entire block as you might expect. For example, if you
        /// translate the range A1:C4 by (2, 2) you will get C4:E6. This transformation should never 
        /// affect the type of cell returned. A ranged cell translated will always be another range. Also note
        /// that you cannot translate past cell A1. Attempting to do so will simply return the translated
        /// range starting at A1.
        /// </remarks>
        public override ICellReference Translate(int rows, int cols)
        {
            // Simply translate the inner points.
            return new RangeCellReference(
                this.StartingCellReference.Translate(rows, cols) as SingleCellReference,
                this.EndingCellReference.Translate(rows, cols) as SingleCellReference);
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Detects if the target single cell reference lies within this range.
        /// </summary>
        /// <param name="cellRef">The cell reference to check.</param>
        /// <returns>True if this range cell reference contains the target single cell reference.</returns>
        private bool Contains(SingleCellReference cellRef)
        {
            // Always dealing with positive numbers.
            return cellRef.ColumnIndex <= this.EndingCellReference.ColumnIndex &&
                cellRef.RowIndex <= this.EndingCellReference.RowIndex;
        }

        /// <summary>
        /// Detects if the target range reference is entirely subsumed by this one.
        /// </summary>
        /// <param name="cellRef">The cell reference to check.</param>
        /// <returns>True if this range subsumes the target, false otherwise.</returns>
        private bool Subsumes(RangeCellReference cellRef)
        {
            return cellRef.StartingCellReference.ColumnIndex <= this.EndingCellReference.ColumnIndex &&
                cellRef.StartingCellReference.RowIndex <= this.EndingCellReference.RowIndex &&
                cellRef.EndingCellReference.ColumnIndex <= this.EndingCellReference.ColumnIndex &&
                cellRef.EndingCellReference.RowIndex <= this.EndingCellReference.RowIndex;
        }

        #endregion
    }
}
