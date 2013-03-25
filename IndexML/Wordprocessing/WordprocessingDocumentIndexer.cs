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

        /// <summary>
        /// A memory stream for the document.
        /// </summary>
        private MemoryStream documentStream;

        #endregion

        #region Constructors & Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="WordprocessingDocumentIndexer "/> class.
        /// </summary>
        /// <param name="toIndex">A stream to the existing document to initialize with. Note that this stream must
        /// have the correct permissions already set on it--that is Read/Write access, in order for the indexer to
        /// be able to access it. Invalid stream modes will throw exceptions when the indexer attempts to create
        /// the spread sheet document.</param>
        /// <exception cref=" ArgumentNullException">Thrown if <paramref name="toIndex"/> is null.</exception>
        /// <exception cref="ArgumentException">Thrown if <paramref name="toIndex"/> is an empty or 
        /// unreadable stream.</exception>
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

        /// <summary>
        /// Initializes a new instance of the <see cref="WordprocessingDocumentIndexer "/> class.
        /// </summary>
        /// <param name="toIndex">The byte array to initialize the indexer with.</param>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="toIndex"/> is null.</exception>
        /// <exception cref="ArgumentException">Thrown if <paramref name="toIndex"/> is empty.</exception>
        public WordprocessingDocumentIndexer(byte[] toIndex)
        {
            if (toIndex == null)
            {
                throw new ArgumentNullException("toIndex");
            }

            if (toIndex.Length == 0)
            {
                throw new ArgumentException("Cannot create an indexer for an empty byte array!", "toIndex");
            }

            var memoryStream = new MemoryStream();
            memoryStream.Write(toIndex, 0, toIndex.Length);
            this.documentStream = memoryStream;
            this.Initialize(WordprocessingDocument.Open(this.Data, true));
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets a value indicating whether this object has been disposed or not.
        /// </summary>
        public bool Disposed { get; private set; }

        /// <summary>
        /// Gets the word processing document the indexer manages. Be careful when making changes to this outside of
        /// using the other indexers; it's easy for them to get out of sync.
        /// </summary>
        public WordprocessingDocument Document { get; private set; }

        /// <summary>
        /// Gets the beginning of the stream containing all the document's information.
        /// </summary>
        public MemoryStream Data
        {
            get
            {
                if (this.Disposed)
                {
                    throw new ObjectDisposedException("WordprocessingDocumentIndexer");
                }

                this.documentStream.Seek(0, SeekOrigin.Begin);
                return this.documentStream;
            }
        }

        /// <summary>
        /// Gets the raw bytes for the spreadsheet document the indexer wraps.
        /// </summary>
        public byte[] Bytes
        {
            get
            {
                if (this.Disposed)
                {
                    throw new ObjectDisposedException("WordprocessingDocumentIndexer");
                }

                return this.Data.ToArray();
            }
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Casts the indexer into a WordprocessingDocument object. Any changes made to the result of this
        /// cast will not be reflected in the indexer, so use this with care.
        /// </summary>
        /// <param name="indexer">The indexer to cast.</param>
        /// <returns>The indexer's wrapped object.</returns>
        public static implicit operator WordprocessingDocument(WordprocessingDocumentIndexer indexer)
        {
            return indexer != null ? indexer.Document : null;
        }

        /// <inheritdoc />
        public void Dispose()
        {
            this.OnDispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Closes the indexer and saves all changes. Also, this call disposes the indexer.
        /// </summary>
        public void SaveAndClose()
        {
            if (!this.Disposed)
            {
                this.Document.Close();
                this.Dispose();
            }
        }

        /// <summary>
        /// Closes the indexer and reopens it. This is a decently heavyweight operation because
        /// it re-indexes the entire document. Use it wisely.
        /// </summary>
        public void SaveAndReopen()
        {
            if (!this.Disposed)
            {
                this.Document.Close();
                this.Initialize(WordprocessingDocument.Open(this.Data, true));
            }
        }

        #endregion

        #region Protected Methods

        /// <summary>
        /// Disposes this object and allows for subclasses to define disposing behavior.
        /// </summary>
        /// <param name="disposing">Has this method been called from the finalizer or the
        /// dispose method?</param>
        protected void OnDispose(bool disposing)
        {
            if (disposing)
            {
                if (this.Document != null)
                {
                    try
                    {
                        this.Document.Close();
                    }
                    catch (Exception)
                    {
                        // Eat it.
                    }
                }

                this.documentStream = null;
                this.Disposed = true;
            }
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Initializes the indexer.
        /// </summary>
        /// <param name="doc">The document to initialize with.</param>
        private void Initialize(WordprocessingDocument doc)
        {
            this.Document = doc;
        }

        #endregion
    }
}
