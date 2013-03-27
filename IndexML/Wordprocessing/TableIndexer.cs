namespace IndexML.Wordprocessing
{
    using System;
    using DocumentFormat.OpenXml.Wordprocessing;

    /// <summary>
    /// OpenXml utilty class for indexing tables.
    /// </summary>
    public class TableIndexer
    {
        private Table table;

        public TableIndexer(Table toIndex)
        {
            if (toIndex == null)
            {
                throw new ArgumentNullException("toIndex");
            }
        }
    }
}
