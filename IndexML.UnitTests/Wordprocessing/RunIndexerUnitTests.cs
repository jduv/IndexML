namespace IndexML.UnitTests.Wordprocessing
{
    using System;
    using DocumentFormat.OpenXml.Wordprocessing;
    using IndexML.Wordprocessing;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [DeploymentItem(@"TestFiles\", @"TestFiles\")]
    public class RunIndexerUnitTests : WordprocessingDocumentTest
    {
        #region Test Methods

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Constructor_NullArgument_ThrowsException()
        {
            var target = new RunIndexer(null);
        }

        [TestMethod]
        public void ImplicitCast_ValidIndexer_SameReference()
        {
            var expected = new Run();
            var indexer = new RunIndexer(expected);
            var target = (Run)indexer;

            Assert.AreSame(expected, indexer.Run);
        }

        [TestMethod]
        public void ImplicitCast_Null_IsNull()
        {
            ParagraphIndexer indexer = null;
            var target = (Paragraph)indexer;
            Assert.IsNull(target);
        }

        #endregion
    }
}
