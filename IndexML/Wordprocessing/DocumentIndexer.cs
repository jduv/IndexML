namespace IndexML.Wordprocessing
{
    using System;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;

    public class DocumentIndexer
    {
        #region Constructors & Destructors

        public DocumentIndexer(MainDocumentPart docPart)
        {
            if (docPart == null)
            {
                throw new ArgumentNullException("doc");
            }
        }

        #endregion

        #region Properties

        public Document Document { get; private set; }

        #endregion
    }
}
