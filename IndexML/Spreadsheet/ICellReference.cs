namespace IndexML
{
    using System;
    using DocumentFormat.OpenXml.Spreadsheet;

    /// <summary>
    /// Defines behavior for a cell reference implementation. A cell reference is a simple wrapper around
    /// a string that allows for simple conversion between string references (A1, B4, etc) and coordinates.
    /// </summary>
    public interface ICellReference
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
        /// Scales the cell reference by the target amounts in each direction.
        /// </summary>
        /// <param name="rows">The number of rows to scale by.</param>
        /// <param name="cols">The number of columns to scale by.</param>
        /// <returns>A scaled cell reference.</returns>
        ICellReference Scale(int rows, int cols);

        /// <summary>
        /// Translates the cell reference by the target amounts in each direction.
        /// </summary>
        /// <param name="rows">The number of rows to translate by.</param>
        /// <param name="cols">The number of columns to translate by.</param>
        /// <returns>A translated cell reference.</returns>
        ICellReference Translate(int rows, int cols);

        #endregion
    }
}
