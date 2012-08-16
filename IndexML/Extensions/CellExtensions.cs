namespace IndexML.Extensions
{
    using DocumentFormat.OpenXml.Spreadsheet;

    /// <summary>
    /// Extension methods for the Cell OpenXml class.
    /// </summary>
    public static class CellExtensions
    {
        #region Extension Methods

        /// <summary>
        /// Determines if the cell has all the proper pieces to it's reference.
        /// </summary>
        /// <param name="toCheck">The cell to check.</param>
        /// <returns>True if the cell has the proper pieces, false otherwise.</returns>
        public static bool HasValidCellRef(this Cell toCheck)
        {
            return toCheck.CellReference != null && !string.IsNullOrEmpty(toCheck.CellReference.Value);
        }

        #endregion
    }
}
