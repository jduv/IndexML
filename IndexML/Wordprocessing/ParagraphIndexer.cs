namespace IndexML.Wordprocessing
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using DocumentFormat.OpenXml.Wordprocessing;

    /// <summary>
    /// OpenXml utiltiy class for indexing paragraph entities.
    /// </summary>
    public class ParagraphIndexer
    {
        #region Fields & Constants

        private IList<RunIndexer> runs = new List<RunIndexer>();

        #endregion

        #region Constructors & Destructors

        public ParagraphIndexer(Paragraph toIndex)
        {
            if (toIndex == null)
            {
                throw new ArgumentNullException("paragraph");
            }

            this.Paragraph = toIndex;
            this.Properties = toIndex.Elements<ParagraphProperties>().FirstOrDefault();

            // Process runs.
            foreach (var run in toIndex.Elements<Run>())
            {
                this.runs.Add(new RunIndexer(run));
            }
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

        /// <summary>
        /// Gets the list of runs inside this paragraph.
        /// </summary>
        public IEnumerable<RunIndexer> Runs
        {
            get
            {
                return this.runs;
            }
        }

        /// <summary>
        /// Gets the text for the paragraph--no markup included.
        /// </summary>
        public string Text
        {
            get
            {
                return this.Paragraph.InnerText;
            }
        }

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
