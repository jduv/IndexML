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

        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentIndexer"/> class.
        /// </summary>
        /// <param name="toIndex">The document part to index.</param>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="toIndex"/> is null.</exception>
        public DocumentIndexer(MainDocumentPart toIndex)
        {
            if (toIndex == null)
            {
                throw new ArgumentNullException("doc");
            }

            this.Document = toIndex.Document;
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

        #endregion
    }
}
