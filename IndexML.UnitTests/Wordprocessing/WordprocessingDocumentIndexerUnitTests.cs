﻿namespace IndexML.UnitTests.Wordprocessing
{
    using System;
    using System.IO;
    using DocumentFormat.OpenXml.Packaging;
    using IndexML.Wordprocessing;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [DeploymentItem(@"TestFiles\", @"TestFiles\")]
    public class WordprocessingDocumentIndexerUnitTests : WordprocessingDocumentTest
    {
        #region Test Methods

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Constructor_NullByteArray_ThrowsException()
        {
            var target = new WordprocessingDocumentIndexer((byte[])null);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void Constructor_EmptyByteArray_ThrowsException()
        {
            var target = new WordprocessingDocumentIndexer(new byte[0]);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Constructor_NullStream_ThrowsException()
        {
            var target = new WordprocessingDocumentIndexer((Stream)null);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void Constructor_EmptyStream_ThrowsException()
        {
            var target = new WordprocessingDocumentIndexer(new MemoryStream());
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void Constructor_UnreadableStream_ThrowsException()
        {
            var stream = new MemoryStream();
            stream.Dispose(); // makes the stream unreadable

            var target = new WordprocessingDocumentIndexer(stream);
        }

        [TestMethod]
        public void Constructor_EmptyDoc_ValidState()
        {
            var docBytes = LoadTestFileBytes(EmptyDocPath);
            if (docBytes != null)
            {
                using (var target = new WordprocessingDocumentIndexer(docBytes))
                {
                    Assert.IsNotNull(target);
                    Assert.IsNotNull(target.Document);
                    Assert.IsFalse(target.Disposed);

                    // Check properties
                    Assert.IsNotNull(target.Data);
                    Assert.IsTrue(target.Data.Length > 0);
                    Assert.IsTrue(target.Bytes.Length > 0);
                }
            }
        }

        [TestMethod]
        public void SaveAndClose_DisposesIndexer()
        {
            AssertFileExists(EmptyDocPath);
            var target = new WordprocessingDocumentIndexer(OpenFileReadWrite(EmptyDocPath));
            target.SaveAndClose();

            Assert.IsTrue(target.Disposed);
        }

        [TestMethod]
        public void SaveAndReopen_DoesNotDisposeIndexer()
        {
            AssertFileExists(EmptyDocPath);
            using (var target = new WordprocessingDocumentIndexer(OpenFileReadWrite(EmptyDocPath)))
            {
                target.SaveAndReopen();
                Assert.IsFalse(target.Disposed);
            }
        }

        [TestMethod]
        [ExpectedException(typeof(ObjectDisposedException))]
        public void Data_DisposedObject()
        {
            AssertFileExists(EmptyDocPath);
            var target = new WordprocessingDocumentIndexer(OpenFileReadWrite(EmptyDocPath));
            target.SaveAndClose();

            Assert.IsTrue(target.Disposed);
            var data = target.Data;
        }

        [TestMethod]
        [ExpectedException(typeof(ObjectDisposedException))]
        public void Bytes_DisposedObject()
        {
            AssertFileExists(EmptyDocPath);
            var target = new WordprocessingDocumentIndexer(OpenFileReadWrite(EmptyDocPath));
            target.SaveAndClose();

            Assert.IsTrue(target.Disposed);
            var data = target.Bytes;
        }

        [TestMethod]
        public void ImplicitCast_ValidIndexer_SameReference()
        {
            AssertFileExists(EmptyDocPath);
            using (var target = new WordprocessingDocumentIndexer(OpenFileReadWrite(EmptyDocPath)))
            {
                WordprocessingDocument doc = (WordprocessingDocument)target;
                Assert.IsNotNull(doc);
                Assert.AreSame(target.WordprocessingDocument, doc);
            }
        }

        [TestMethod]
        public void ImplicitCast_Null_IsNull()
        {
            WordprocessingDocumentIndexer indexer = null;
            var target = (WordprocessingDocument)indexer;
            Assert.IsNull(target);
        }

        #endregion
    }
}
