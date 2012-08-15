namespace IndexML
{
    using System;
    using DocumentFormat.OpenXml.Spreadsheet;

    /// <summary>
    /// Defines behavior for a cell reference implementation. A cell reference is a simple wrapper around
    /// a string that allows for simple conversion between string references (A1, B4, etc) and coordinates.
    /// </summary>
    public interface ICellReference : IEquatable<ICellReference>
    {
        #region Interface Properties

        /// <summary>
        /// Gets the string value of the cell reference.
        /// </summary>
        string Value { get; }

        #endregion

        #region Interface Methods

        /// <summary>
        /// Does the cell reference contain the other cell reference?
        /// </summary>
        /// <param name="cellRef">The cell reference to check for.</param>
        /// <returns>True if the reference contains the target cell, false otherwise.</returns>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="cellRef"/> is null.</exception>
        bool ContainsOrSubsumes(ICellReference cellRef);

        /// <summary>
        ///  Extends the column range of this cell reference by the target length and returns
        ///  the result.
        /// </summary>
        /// <param name="length">The number of columns to extend the range to.</param>
        /// <returns>A new cell reference implementation with an extended range.</returns>
        ICellReference ExtendColumnRange(int length);

        /// <summary>
        ///  Extends the row range of this cell reference by the target length and returns
        ///  the result.
        /// </summary>
        /// <param name="length">The number of columns to extend the range to.</param>
        /// <returns>A new cell reference implementation with an extended range.</returns>
        ICellReference ExtendRowRange(int length);

        #endregion
    }
}
