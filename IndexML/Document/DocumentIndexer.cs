namespace IndexML.Document
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using DocumentFormat.OpenXml.Wordprocessing;

    /// <summary>
    /// OpenXML utility class for Indexing word documents.
    /// </summary>
    public class DocumentIndexer
    {

        #region Constructors & Destructors

        public DocumentIndexer(DocPart toIndex)
        {
        }

        #endregion

        #region Properties

        public Document Document { get; private set; }

        #endregion
    }
}
