namespace IndexML.Wordprocessing
{
    using System;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;

    /// <summary>
    /// OpenXml utility class for managing a document.
    /// </summary>
    public class DocumentIndexer
    {
        #region Constructors & Destructors

        public DocumentIndexer(Document toIndex)
        {
            if (toIndex == null)
            {
                throw new ArgumentNullException("doc");
            }

            this.Document = toIndex;
            this.Body = new BodyIndexer(this.Document.Body);
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets the document.
        /// </summary>
        public Document Document { get; private set; }

        /// <summary>
        /// Gets the body indexer.
        /// </summary>
        public BodyIndexer Body { get; private set; }

        /// <summary>
        /// Gets the numbering indexer
        /// </summary>
        public NumberingIndexer Numbering { get; private set; }

        #endregion

        #region Public Methods

        /// <summary>
        /// Casts the indexer to a Document element. Any changes made to the raw element will not
        /// be reflected in the indexer.
        /// </summary>
        /// <param name="indexer">The indexer to cast.</param>
        /// <returns>The element that the indexer wraps, or null if <paramref name="indexer"/> is null.</returns>
        public static implicit operator Document(DocumentIndexer indexer)
        {
            return indexer != null ? indexer.Document : null;
        }

        #endregion
    }
}
