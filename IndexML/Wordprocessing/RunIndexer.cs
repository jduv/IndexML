namespace IndexML.Wordprocessing
{
    using System;
    using DocumentFormat.OpenXml.Wordprocessing;

    /// <summary>
    /// OpenXml utility class for performing operations on text runs.
    /// </summary>
    public class RunIndexer
    {
        #region Constructors & Destructors

        public RunIndexer(Run toIndex)
        {
            if (toIndex == null)
            {
                throw new ArgumentNullException("toIndex");
            }

            this.Run = toIndex;
            this.Properties = toIndex.RunProperties;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets the contained run.
        /// </summary>
        public Run Run { get; private set; }

        /// <summary>
        /// Gets the properties for the run.
        /// </summary>
        public RunProperties Properties { get; private set; }

        /// <summary>
        /// Gets a value indicating whether this run is bolded.
        /// </summary>
        public bool IsBold
        {
            get
            {
                return this.Properties != null && this.Properties.Bold != null;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this run is italic.
        /// </summary>
        public bool IsItalic
        {
            get
            {
                return this.Properties != null && this.Properties.Italic != null;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this run is emphatic.
        /// </summary>
        public bool IsEmphatic
        {
            get
            {
                return this.Properties != null && this.Properties.Emphasis != null;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this run is underline.
        /// </summary>
        public bool IsUnderline
        {
            get
            {
                return this.Properties != null && this.Properties.Underline != null;
            }
        }

        /// <summary>
        /// Gets the text for the run.
        /// </summary>
        public string Text
        {
            get
            {
                return this.Run.InnerText;
            }
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Casts the indexer to a Run element. Any changes made to the raw element will not
        /// be reflected in the indexer.
        /// </summary>
        /// <param name="indexer">The indexer to cast.</param>
        /// <returns>The element that the indexer wraps, or null if <paramref name="indexer"/> is null.</returns>
        public static implicit operator Run(RunIndexer indexer)
        {
            return indexer != null ? indexer.Run : null;
        }

        #endregion
    }
}
