namespace IndexML.UnitTests.Wordprocessing
{
    using System;
    using System.Linq;
    using DocumentFormat.OpenXml.Wordprocessing;
    using IndexML.Wordprocessing;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [DeploymentItem(@"TestFiles\", @"TestFiles\")]
    public class BodyIndexerUnitTests : WordprocessingDocumentTest
    {
        #region Test Methods

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Constructor_NullArgument_ThrowsException()
        {
            var target = new BodyIndexer(null);
        }

        [TestMethod]
        public void Constructor_EmptyDocument_ValidState()
        {
            SafeExecuteTest(
                EmptyDocPath,
                (doc) =>
                {
                    var target = new BodyIndexer(doc.MainDocumentPart.Document.Body);

                    Assert.IsNotNull(target.Paragraphs);
                    Assert.AreEqual(1, target.Paragraphs.Count()); // Only one paragraph in an empty doc.
                    Assert.IsNotNull(target.Tables);
                    Assert.AreEqual(0, target.Tables.Count()); // No tables.
                    Assert.IsNotNull(target.FinalSectionProperties);

                    // Check references.
                    Assert.AreSame(
                        doc.MainDocumentPart.Document.Body.Elements<SectionProperties>().FirstOrDefault(), 
                        target.FinalSectionProperties);
                });
        }

        [TestMethod]
        public void Constructor_StandardDocument_ValidState()
        {
            SafeExecuteTest(
                StandardDocPath,
                (doc) =>
                {
                    var target = new BodyIndexer(doc.MainDocumentPart.Document.Body);

                    Assert.IsNotNull(target.Paragraphs);
                    Assert.IsTrue(target.Paragraphs.Count() > 1);
                    Assert.IsNotNull(target.Tables);
                    Assert.IsTrue(target.Tables.Count() > 0);
                    Assert.IsNotNull(target.FinalSectionProperties);

                    // Check references.
                    Assert.AreSame(
                        doc.MainDocumentPart.Document.Body.Elements<SectionProperties>().FirstOrDefault(),
                        target.FinalSectionProperties);
                });
        }

        [TestMethod]
        public void ImplicitCast_ValidIndexer_SameReference()
        {
            var expected = new Body();
            var indexer = new BodyIndexer(expected);
            var target = (Body)indexer;

            Assert.AreSame(expected, target);
        }

        [TestMethod]
        public void ImplicitCast_Null_IsNull()
        {
            BodyIndexer indexer = null;
            var target = (Body)indexer;
            Assert.IsNull(target);
        }

        #endregion
    }
}
