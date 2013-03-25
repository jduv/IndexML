namespace IndexML.Wordprocessing
{
    using System;
    using System.IO;
    using DocumentFormat.OpenXml.Packaging;
    using IndexML.Extensions;

    /// <summary>
    /// OpenXml utility class for performing common operations on a word processing document.
    /// </summary>
    public class WordprocessingDocumentIndexer : IDisposable
    {
        #region Fields & Constants

        private MemoryStream documentStream;

        #endregion

        #region Constructors & Destructors

        public WordprocessingDocumentIndexer(Stream toIndex)
        {
            if (toIndex == null)
            {
                throw new ArgumentNullException("toIndex");
            }

            if (!toIndex.CanRead || toIndex.Length <= 0 || !toIndex.CanSeek)
            {
                throw new ArgumentException("Cannot create an indexer for a document with an empty or unreadable stream!", "toIndex");
            }

            this.documentStream = new MemoryStream();
            StreamExtensions.Copy(toIndex, this.documentStream);
            this.Initialize(WordprocessingDocument.Open(this.Data, true));
        } 

        #endregion

        #region Properties

        public WordprocessingDocument Document { get; private set; }

        public MemoryStream Data
        {
            get
            {
                this.documentStream.Seek(0, SeekOrigin.Begin);
                return this.documentStream;
            }
        }

        #endregion

        #region Public Methods

        public static implicit operator WordprocessingDocument(WordprocessingDocumentIndexer indexer)
        {
            return indexer != null ? indexer.Document : null;
        }

        #endregion

        #region Private Methods

        private void Initialize(WordprocessingDocument doc)
        {
            this.Document = doc;
        }

        #endregion
    }
}
