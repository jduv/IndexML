namespace IndexML.Spreadsheet
{
    using System;
    using System.IO;
    using DocumentFormat.OpenXml.Packaging;

    /// <summary>
    /// OpenXml utility class for performing common operations on a spreadsheet document including
    /// random access support for rows, inserting/adding new rows, deleting rows, and replacing text using
    /// the shared string model.
    /// </summary>
    public class SpreadsheetIndexer : IDisposable
    {
        #region Fields & Constants

        /// <summary>
        /// The document stream for the spread sheet.
        /// </summary>
        private MemoryStream documentStream;

        #endregion

        #region Constructors & Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="SpreadsheetIndexer"/> class.
        /// </summary>
        /// <param name="toIndex">A stream to the existing document to initialize with. Note that this stream must
        /// have the correct permissions already set on it--that is Read/Write access, in order for the indexer to
        /// be able to access it. Invalid stream modes will throw exceptions when the indexer attempts to create
        /// the spread sheet document.</param>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="toIndex"/> is null.</exception>
        /// <exception cref="ArgumentException">Thrown if <paramref name="toIndex"/> is an empty or
        /// unreadable stream.</exception>
        public SpreadsheetIndexer(Stream toIndex)
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
            CopyStream(toIndex, this.documentStream);
            
            this.Initialize(SpreadsheetDocument.Open(this.Data, true));
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SpreadsheetIndexer"/> class.
        /// </summary>
        /// <param name="toIndex">The spreadsheet document to initialize the indexer with.</param>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="toIndex"/> is null.</exception>
        public SpreadsheetIndexer(byte[] toIndex)            
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
            this.Initialize(SpreadsheetDocument.Open(this.Data, true));
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets the Spreadsheet that the indexer manages. Be careful when making changes to this outside of
        /// using the other indexers; it's easy for them to get out of sync.
        /// </summary>
        public SpreadsheetDocument Spreadsheet { get; private set; }

        /// <summary>
        /// Gets a <see cref="WorkbookIndexer"/> for the document's workbook.
        /// </summary>
        public WorkbookIndexer Workbook { get; private set; }

        /// <summary>
        /// Gets the beginning of the stream containing all the document's information.
        /// </summary>
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

        /// <summary>
        /// Casts the indexer into a Spreadsheet object. Any changes made to the result of this
        /// cast will not be reflected in the indexer, so use this with care.
        /// </summary>
        /// <param name="indexer">The indexer to cast.</param>
        /// <returns>The indexer's wrapped object.</returns>
        public static implicit operator SpreadsheetDocument(SpreadsheetIndexer indexer)
        {
            return indexer != null ? indexer.Spreadsheet : null;
        }

        /// <inheritdoc />
        public void Dispose()
        {
            this.OnDispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Gets the bytes of this indexer.
        /// </summary>
        /// <returns>A byte array of the data inside this indexer.</returns>
        public byte[] GetBytes()
        {
            return this.Data.ToArray();
        }

        /// <summary>
        /// Closes the spreadsheet indexer.
        /// </summary>
        public void SaveAndClose()
        {
            this.Spreadsheet.Close();
        }

        /// <summary>
        /// Closes the indexer and reopens it. This is a decently heavyweight operation because
        /// it re-indexes the entire document. Use it wisely.
        /// </summary>
        public void SaveAndReopen()
        {
            this.Spreadsheet.Close();
            this.Initialize(SpreadsheetDocument.Open(this.Data, true));
        }

        #endregion

        #region Protected Methods

        /// <summary>
        /// Disposes this object, and allows for subclasses to define disposing behavior.
        /// </summary>
        /// <param name="disposing">Has this method been called from the finalizer or the
        /// dispose method?</param>
        protected virtual void OnDispose(bool disposing)
        {
            // If we're disposing, then we've likely got all our handles.
            if (disposing)
            {
                if (this.Spreadsheet != null)
                {
                    try
                    {
                        this.Spreadsheet.Close();
                    }
                    catch (Exception)
                    {
                        // Eat it.
                    }
                }              
            }
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Copies the source stream into the target.
        /// </summary>
        /// <param name="source">The source stream.</param>
        /// <param name="target">The target stream.</param>
        private static void CopyStream(Stream source, Stream target)
        {
            using (source)
            {
                var buffer = new byte[32768];
                int bytesRead;
                while ((bytesRead = source.Read(buffer, 0, buffer.Length)) > 0)
                {
                    target.Write(buffer, 0, bytesRead);
                }
            }
        }

        /// <summary>
        /// Initializes the indexer.
        /// </summary>
        /// <param name="spreadsheet">The spreadsheet to initialize with.</param>
        private void Initialize(SpreadsheetDocument spreadsheet)
        {
            this.Spreadsheet = spreadsheet;
            this.Workbook = new WorkbookIndexer(spreadsheet.WorkbookPart);
        }

        #endregion
    }
}
