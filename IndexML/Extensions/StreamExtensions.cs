namespace IndexML.Extensions
{
    using System.IO;

    /// <summary>
    /// Holds extension and utility methods for streams.
    /// </summary>
    public static class StreamExtensions
    {
        #region Public Methods

        /// <summary>
        /// Copies the source stream into the target.
        /// </summary>
        /// <param name="source">The source stream.</param>
        /// <param name="target">The target stream.</param>
        public static void Copy(Stream source, Stream target)
        {
            using (source)
            {
                var buffer = new byte[32768];
                int bytesRead;
                while ((bytesRead = source.Read(buffer, 0, buffer.Length)) > 0)
                {
                    target.Write(buffer, 0, bytesRead);
                }
            }
        }

        #endregion
    }
}
