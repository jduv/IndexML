namespace IndexML.Wordprocessing
{
    using System;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;

    /// <summary>
    /// OpenXml utility class for handling Numbering objects in a word processing document.
    /// This is required in order to handle specific list scenarios when parsing run texts.
    /// </summary>
    public class NumberingIndexer
    {
        #region Constructors & Destructors

        public NumberingIndexer(NumberingDefinitionsPart toIndex)
        {
            if (toIndex == null)
            {
                throw new ArgumentNullException("toIndex");
            }

            this.Numbering = toIndex.Numbering;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets the wrapped numbering element.
        /// </summary>
        public Numbering Numbering { get; private set; }

        #endregion

        #region Public Methods

        /// <summary>
        /// Casts the indexer to a Numbering element. Any changes made to the raw element will not
        /// be reflected in the indexer.
        /// </summary>
        /// <param name="indexer">The indexer to cast.</param>
        /// <returns>The element that the indexer wraps, or null if <paramref name="indexer"/> is null.</returns>
        public static implicit operator Numbering(NumberingIndexer indexer)
        {
            return indexer != null ? indexer.Numbering : null;
        }

        #endregion
    }
}
