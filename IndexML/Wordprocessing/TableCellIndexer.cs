namespace IndexML.Wordprocessing
{
    using System;
    using System.Collections.Generic;
    using DocumentFormat.OpenXml.Wordprocessing;

    /// <summary>
    /// OpenXml utility class for manipulating table cells.
    /// </summary>
    public class TableCellIndexer
    {
        #region Fields & Constants

        private readonly IList<TableIndexer> childTables;
        private readonly IList<ParagraphIndexer> paragraphs;

        #endregion

        #region Constructors & Destructors

        public TableCellIndexer(TableCell toIndex)
        {
            if (toIndex == null)
            {
                throw new ArgumentNullException("toIndex");
            }

            this.Cell = toIndex;
            this.childTables = new List<TableIndexer>();
            this.paragraphs = new List<ParagraphIndexer>();

            // Again, we forgo linq here for performance reasons. One time
            // through the loop should be enough.
            foreach (var element in toIndex.Elements())
            {
                if (element != null)
                {
                    if (element is Paragraph)
                    {
                        this.paragraphs.Add(new ParagraphIndexer(element as Paragraph));
                    }
                    else if (element is Table)
                    {
                        // This could get nasty with lots of nested stuff.
                        this.childTables.Add(new TableIndexer(element as Table));
                    }
                }
            }
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets the wrapped cell.
        /// </summary>
        public TableCell Cell { get; private set; }

        /// <summary>
        /// Gets an enumeration of the paragraphs inside the cell.
        /// </summary>
        public IEnumerable<ParagraphIndexer> Paragraphs
        {
            get
            {
                return this.paragraphs;
            }
        }

        /// <summary>
        /// Gets an enumeration of the child tables inside the cell
        /// </summary>
        public IEnumerable<TableIndexer> Tables
        {
            get
            {
                return this.childTables;
            }
        }

        /// <summary>
        /// Gets the text for the cell. This simply concatenates the text of all enclosed
        /// paragraphs. Tables are ignored.
        /// </summary>
        public string Text
        {
            get
            {
                string text = string.Empty;
                foreach (var p in this.Paragraphs)
                {
                    text += p.Text;
                }

                return text;
            }
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Casts the indexer to a TableCell element. Any changes made to the raw element will not
        /// be reflected in the indexer.
        /// </summary>
        /// <param name="indexer">The indexer to cast.</param>
        /// <returns>The element that the indexer wraps, or null if <paramref name="indexer"/> is null.
        /// </returns>
        public static implicit operator TableCell(TableCellIndexer indexer)
        {
            return indexer != null ? indexer.Cell : null;
        }

        #endregion
    }
}
