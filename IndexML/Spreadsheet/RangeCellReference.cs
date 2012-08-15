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
