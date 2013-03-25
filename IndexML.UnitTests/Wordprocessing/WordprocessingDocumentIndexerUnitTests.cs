﻿namespace IndexML.UnitTests.Wordprocessing
{
    using System;
    using System.IO;
    using IndexML.Wordprocessing;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    public class WordprocessingDocumentIndexerUnitTests : WordprocessingDocumentTest
    {
        #region Test Methods

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Constructor_NullArgument_ThrowsException()
        {
            var target = new WordprocessingDocumentIndexer((byte[])null);
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
        [DeploymentItem(EmptyDocPath, TestFilesDir)]
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
                }
            }
        }

        [TestMethod]
        [DeploymentItem(EmptyDocPath, TestFilesDir)]
        public void SaveAndClose_DisposesIndexer()
        {
            AssertFileExists(EmptyDocPath);
            var target = new WordprocessingDocumentIndexer(
                File.Open(EmptyDocPath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite));
            target.SaveAndClose();

            Assert.IsTrue(target.Disposed);
        }

        [TestMethod]
        [DeploymentItem(EmptyDocPath, TestFilesDir)]
        public void SaveAndReopen_DoesNotDisposeIndexer()
        {
            AssertFileExists(EmptyDocPath);
            var target = new WordprocessingDocumentIndexer(
                File.Open(EmptyDocPath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite));
            target.SaveAndReopen();

            Assert.IsFalse(target.Disposed);
        }

        [TestMethod]
        [DeploymentItem(EmptyDocPath, TestFilesDir)]
        [ExpectedException(typeof(ObjectDisposedException))]
        public void DataProperty_DisposedObject()
        {
            AssertFileExists(EmptyDocPath);
            var target = new WordprocessingDocumentIndexer(
                File.Open(EmptyDocPath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite));
            target.SaveAndClose();

            Assert.IsTrue(target.Disposed);
            var data = target.Data;
        }

        [TestMethod]
        [DeploymentItem(EmptyDocPath, TestFilesDir)]
        [ExpectedException(typeof(ObjectDisposedException))]
        public void BytesProperty_DisposedObject()
        {
            AssertFileExists(EmptyDocPath);
            var target = new WordprocessingDocumentIndexer(
                File.Open(EmptyDocPath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite));
            target.SaveAndClose();

            Assert.IsTrue(target.Disposed);
            var data = target.Bytes;
        }

        #endregion
    }
}