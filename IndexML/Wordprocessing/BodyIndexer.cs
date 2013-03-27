namespace IndexML.Wordprocessing
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using DocumentFormat.OpenXml.Wordprocessing;

    /// <summary>
    /// OpenXml utility class for indexing word document bodies.
    /// </summary>
    public class BodyIndexer
    {
        #region Fields & Constants

        /// <summary>
        /// The list of paragraphs inside the document.
        /// </summary>
        private IList<ParagraphIndexer> paragraphs;

        /// <summary>
        /// The list of tables inside the document.
        /// </summary>
        private IList<TableIndexer> tables;

        #endregion

        #region Constructors & Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="BodyIndexer"/> class.
        /// </summary>
        /// <param name="toIndex">The document body to index.</param>
        public BodyIndexer(Body toIndex)
        {
            if (toIndex == null)
            {
                throw new ArgumentNullException("toIndex");
            }

            this.paragraphs = new List<ParagraphIndexer>();
            this.tables = new List<TableIndexer>();

            // This is uglier than Linq, but way more efficient. For large documents, we only
            // wish to go throught the entire list fo elements ONCE, not multiple times.
            foreach (var element in toIndex.Elements())
            {
                if (element != null)
                {
                    if (element is Paragraph)
                    {
                        var indexer = new ParagraphIndexer(element as Paragraph);
                        this.paragraphs.Add(indexer);
                    }
                    else if (element is Table)
                    {
                        var indexer = new TableIndexer(element as Table);
                        this.tables.Add(indexer);
                    }
                    else if (element is SectionProperties)
                    {
                        this.FinalSectionProperties = element as SectionProperties;
                    }
                }
            }

            // This element should always exist.
            this.FinalSectionProperties = toIndex.Elements<SectionProperties>().FirstOrDefault();
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets the final section properties for the body.
        /// </summary>
        public SectionProperties FinalSectionProperties { get; private set; }

        /// <summary>
        /// Gets a list of paragraphs in the document.
        /// </summary>
        public IEnumerable<ParagraphIndexer> Paragraphs
        {
            get
            {
                return this.paragraphs;
            }
        }

        /// <summary>
        /// Gets a list of Tables in the document.
        /// </summary>
        public IEnumerable<TableIndexer> Tables
        {
            get
            {
                return this.tables;
            }
        }

        #endregion
    }
}
