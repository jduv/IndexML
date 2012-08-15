namespace IndexML
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;

    /// <summary>
    /// OpenXml utility class for performing operations on shared string table objects.
    /// </summary>
    public class SharedStringTableIndexer
    {
        #region Fields & Constants

        /// <summary>
        /// Internal index for keeping up with where we are in the string table.
        /// </summary>
        private long currentIndex = 0;

        /// <summary>
        /// Holds all the indices mapped from their corresponding strings.
        /// </summary>
        private IDictionary<string, long> indexDict = new Dictionary<string, long>();

        /// <summary>
        /// Holds all the strings mapped from their corresponding indices.
        /// </summary>
        private IDictionary<long, string> strDict = new Dictionary<long, string>();

        #endregion

        #region Constructors & Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="SharedStringTableIndexer"/> class. If the 
        /// target shared string table part doesn't have a string table initialized, one will be created 
        /// and appended to the part.
        /// </summary>
        /// <param name="toIndex">The shared string table part instance to index.</param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="toIndex"/> is null.
        /// </exception>
        public SharedStringTableIndexer(SharedStringTablePart toIndex)
        {
            // Handle parameter checking in ctor.
            if (toIndex == null)
            {
                throw new ArgumentNullException("stringTablePart");
            }

            // In this case, we'll initialize the shared string table object for the caller.
            if (toIndex.SharedStringTable == null)
            {
                // Side-effect!
                toIndex.SharedStringTable = new SharedStringTable()
                {
                    Count = 0U,
                    UniqueCount = 0U
                };
            }
            
            this.Initialize(toIndex.SharedStringTable);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SharedStringTableIndexer"/> class.
        /// </summary>
        /// <param name="toIndex">The string table to index.</param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="toIndex"/> is null.
        /// </exception>
        public SharedStringTableIndexer(SharedStringTable toIndex)
        {
            // Handle parameter checking in ctor.
            if (toIndex == null)
            {
                throw new ArgumentNullException("toIndex");
            }
            
            this.Initialize(toIndex);
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets the object associated with this indexer. Changes made to the cell will not be reflected
        /// to any dependent properties inside the indexer, so use this with care.
        /// </summary>
        public SharedStringTable SharedStringTable { get; private set; }

        /// <summary>
        /// Gets the shared string items inside the indexed shared string table.
        /// </summary>
        public IEnumerable<SharedStringItem> Items
        {
            get
            {
                return this.SharedStringTable.Descendants<SharedStringItem>();
            }
        }

        /// <summary>
        /// Gets the number of shared string items in this indexer.
        /// </summary>
        public long UniqueCount
        {
            get
            {
                return this.SharedStringTable.UniqueCount;
            }
        }

        /// <summary>
        /// Gets the number of shared strings in the string table--this count is not unique.
        /// </summary>
        public long Count
        {
            get
            {
                return this.SharedStringTable.Count;
            }
        }

        /// <summary>
        /// Gets the string value at the target index.
        /// </summary>
        /// <param name="idx">The index to retrieve.</param>
        /// <returns>The string value at the target index.</returns>
        /// <exception cref="KeyNotFoundException">Thrown when the index doesn't
        /// exist in the indexer's index dictionary.</exception>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="idx"/> is null.
        /// </exception>
        public string this[long idx]
        {
            get
            {
                // Let standard dictionary exceptions bubble up.
                return this.strDict[idx];
            }
        }

        /// <summary>
        /// Gets the index value for the given string.
        /// </summary>
        /// <param name="str">The string whose index to retrieve.</param>
        /// <returns>The index of the target string.</returns>
        /// <exception cref="KeyNotFoundException">Thrown when the string doesn't
        /// exist in the indexer's string dictionary.</exception>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="str"/> is null.
        /// </exception>
        public long this[string str]
        {
            get
            {
                // Let standard dictionary exceptions bubble up.
                return this.indexDict[str];
            }
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Casts the indexer into a SharedStringTable object. Any changes made to the result of this
        /// cast will not be reflected in the indexer, so use this with care.
        /// </summary>
        /// <param name="indexer">The indexer to cast.</param>
        /// <returns>The indexer's wrapped object.</returns>
        public static implicit operator SharedStringTable(SharedStringTableIndexer indexer)
        {
            return indexer != null ? indexer.SharedStringTable : null;
        }

        /// <summary>
        /// Adds the target string to the indexer. If the string is already detected to exist inside
        /// the indexer, then it will not be duplicated.
        /// </summary>
        /// <param name="toAdd">The string to add.</param>
        /// <param name="preserveSpace">Should spaces be preserved? By default this is set to
        /// false.</param>
        /// <returns>The index of the string inside the shared string table.</returns>
        public long Add(string toAdd, bool preserveSpace = false)
        {
            if (!this.indexDict.ContainsKey(toAdd))
            {
                var sharedStrItem = new SharedStringItem()
                {
                    Text = new Text()
                    {
                        Space = preserveSpace ?
                            SpaceProcessingModeValues.Preserve :
                            SpaceProcessingModeValues.Default,
                        Text = toAdd
                    }
                };

                this.SharedStringTable.Append(sharedStrItem);

                // now update our books
                this.indexDict[toAdd] = this.currentIndex;
                this.strDict[this.currentIndex] = toAdd;
                this.SharedStringTable.Count++;
                this.SharedStringTable.UniqueCount++;

                return this.currentIndex++;
            }
            else
            {
                return this.indexDict[toAdd];
            }
        }

        /// <summary>
        /// Adds a list of strings to the indexer. Duplicates will be ignored.
        /// </summary>
        /// <param name="toAdd">The list of strings to add.</param>
        /// <returns>A list of indices corresponding to the list of strings that were 
        /// added ot the indexer. Order is preserved.</returns>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="toAdd"/> is null.</exception>
        public IList<long> AddAll(IEnumerable<string> toAdd)
        {
            if (toAdd == null)
            {
                throw new ArgumentNullException("toAdd");
            }

            // Allocate the memory. Defered execution is not appropriate here.
            var idxList = new List<long>();
            foreach (var s in toAdd)
            {
                idxList.Add(this.Add(s));
            }

            return idxList;
        }

        /// <summary>
        /// Does the indexer contain the target string?
        /// </summary>
        /// <param name="str">The shared string to check for.</param>
        /// <returns>True if the indexer contains the target shared string, false otherwise.</returns>
        public bool Contains(string str)
        {
            return this.indexDict.ContainsKey(str);
        }

        /// <summary>
        /// Does the indexer contain a string at the target index?
        /// </summary>
        /// <param name="idx">The index to check for.</param>
        /// <returns>True if the indexer contains a shared string entry for the target index,
        /// false otherwise.</returns>
        public bool Contains(long idx)
        {
            return this.strDict.ContainsKey(idx);
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Initializes the indexer from the target shared string table.
        /// </summary>
        /// <param name="toIndex">The shared string table to index.</param>
        private void Initialize(SharedStringTable toIndex)
        {
            this.SharedStringTable = toIndex;

            if (toIndex.Count > 0)
            {
                // Table already has entries
                var sharedStrings = toIndex.Descendants<SharedStringItem>()
                    .Select(x => x.InnerText);

                long idx = 0;
                foreach (var sharedStr in sharedStrings)
                {
                    this.indexDict[sharedStr] = idx;
                    this.strDict[idx] = sharedStr;
                    idx++;
                }

                this.currentIndex = toIndex.UniqueCount;
            }
        }

        #endregion
    }
}
