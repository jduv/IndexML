namespace IndexML.Extensions
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    public static class IEnumerableExtensions
    {
        #region Extension Methods

        /// <summary>
        /// A simple, optimized version of FirstOrDefault. This method will spin through the target
        /// enumerable looking for an item that matches the target predicate. If the predicate fails,
        /// it immediately returns default(T). The benefits may seem subtle, but in essence this allows
        /// you to look for items based on specific criteria and not spin through the entire list when
        /// you *know* that the item doesn't exist based on some implicit criteria.
        /// </summary>
        /// <typeparam name="T">The type of objects inside the enumerable.</typeparam>
        /// <param name="source">The source enumerable.</param>
        /// <param name="terminator">A terminating condition.</param>
        /// <returns>The target value, or null if not found.</returns>
        public static T FindOrDefault<T>(this IEnumerable<T> source, Predicate<T> terminator)
        {
            if (terminator == null)
            {
                throw new ArgumentNullException("terminator");
            }

            T target = default(T);
            foreach (var item in source)
            {
                if (terminator(item))
                {
                    target = item;
                    break;
                }
            }

            return target;
        }

        #endregion
    }
}
