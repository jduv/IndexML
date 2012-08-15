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
        /// Does teh cell reference contain the other cell reference?
        /// </summary>
        /// <param name="cellRef">The cell reference to check for.</param>
        /// <returns>True if the reference contains the target cell, false otherwise.</returns>
        bool ContainsOrSubsumes(ICellReference cellRef);

        #endregion
    }
}
