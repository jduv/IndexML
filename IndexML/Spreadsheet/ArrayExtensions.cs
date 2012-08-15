namespace IndexML
{
    /// <summary>
    /// Extensions to the Array class.
    /// </summary>
    public static class ArrayExtensions
    {
        #region Extension Methods

        /// <summary>
        /// Swaps the item at the first index with the item at the second.
        /// </summary>
        /// <typeparam name="T">The type of objects contained within the array.</typeparam>
        /// <param name="array">The array to perform the swap on.</param>
        /// <param name="index1">The first index.</param>
        /// <param name="index2">The second index.</param>
        public static void Swap<T>(this T[] array, long index1, long index2)
        {
            var temp = array[index2];
            array[index2] = array[index1];
            array[index1] = temp;
        }

        #endregion
    }
}
