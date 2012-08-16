namespace IndexML.Extensions
{
    /// <summary>
    /// Extensions for the string class.
    /// </summary>
    public static class StringExtensions
    {
        #region Extension Methods

        /// <summary>
        /// A simple method that pretty prints a string if it's null or empty. Mainly used
        /// for exception throws and debug output.
        /// </summary>
        /// <param name="str">This string.</param>
        /// <returns>The pretty print version of this string.</returns>
        public static string PrettyPrint(this string str)
        {
            if (str == null)
            {
                return "<null>";
            }
            else if (str == string.Empty)
            {
                return "<empty>";
            }
            else
            {
                return str;
            }
        }

        #endregion
    }
}
