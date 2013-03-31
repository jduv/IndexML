namespace IndexML.Wordprocessing
{
    using System;
    using System.Linq;
    using DocumentFormat.OpenXml.Wordprocessing;

    /// <summary>
    /// OpenXml utiltiy class for indexing paragraph entities.
    /// </summary>
    public class ParagraphIndexer
    {
        #region Constructors & Destructors

        public ParagraphIndexer(Paragraph toIndex)
        {
            if (toIndex == null)
            {
                throw new ArgumentNullException("paragraph");
            }

            this.Paragraph = toIndex;
            this.Properties = toIndex.Elements<ParagraphProperties>().FirstOrDefault();
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets the wrapped paragraph.
        /// </summary>
        public Paragraph Paragraph { get; private set; }

        /// <summary>
        /// Gets the paragraph's properties.
        /// </summary>
        public ParagraphProperties Properties { get; private set; }

        #endregion

        #region Public Methods

        /// <summary>
        /// Casts the indexer to a Paragraph element. Any changes made to the raw element will not
        /// be reflected in the indexer.
        /// </summary>
        /// <param name="indexer">The indexer to cast.</param>
        /// <returns>The element that the indexer wraps, or null if <paramref name="indexer"/> is null.</returns>
        public static implicit operator Paragraph(ParagraphIndexer indexer)
        {
            return indexer != null ? indexer.Paragraph : null;
        }

        #endregion
    }
}
