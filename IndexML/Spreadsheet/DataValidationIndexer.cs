﻿namespace IndexML
{
    using System;
    using System.Collections.Generic;
    using DocumentFormat.OpenXml.Spreadsheet;

    /// <summary>
    /// An indexer for a data validation OpenXml element.
    /// </summary>
    public sealed class DataValidationIndexer
    {
        #region Fields & Constants

        /// <summary>
        /// A hash set of references.
        /// </summary>
        private readonly IList<ICellReference> cellReferences = new List<ICellReference>();

        #endregion

        #region Constructors & Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="DataValidationIndexer"/> class.
        /// </summary>
        /// <param name="toIndex">The validator to index.</param>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="toIndex"/> is null.</exception>
        public DataValidationIndexer(DataValidation toIndex)
        {
            if (toIndex == null)
            {
                throw new ArgumentNullException("toIndex");
            }

            this.DataValidation = toIndex;

            foreach (var reference in toIndex.SequenceOfReferences.Items)
            {
                this.cellReferences.Add(CellReference.Create(reference.Value));
            }
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets the data validation element the indexer manages.
        /// </summary>
        public DataValidation DataValidation { get; private set; }

        /// <summary>
        /// Gets an enumerable of cell reference strings.
        /// </summary>
        public IEnumerable<ICellReference> CellReferences
        {
            get
            {
                foreach (var item in this.cellReferences)
                {
                    yield return item;
                }
            }
        }

        /// <summary>
        /// Gets the number of references inside the indexer.
        /// </summary>
        public int ReferenceCount
        {
            get
            {
                return this.cellReferences.Count;
            }
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Casts the indexer into a DataValidation object. Any changes made to the result of this
        /// cast will not be reflected in the indexer, so use this with care.
        /// </summary>
        /// <param name="toCast">The indexer to cast.</param>
        /// <returns>The indexer's wrapped object.</returns>
        public static implicit operator DataValidation(DataValidationIndexer toCast)
        {
            return toCast != null ? toCast.DataValidation : null;
        }

        /// <summary>
        /// Adds a cell reference to the indexer, only if the reference is valid. Ignores null.
        /// </summary>
        /// <param name="toAdd">The cell to add.</param>        
        public void Add(Cell toAdd)
        {
            if (toAdd.HasValidCellRef())
            {
                // Only add the new cell reference if it doesn't already exist.
                var refToAdd = CellReference.Create(toAdd.CellReference.Value);
                if (!this.cellReferences.Contains(refToAdd))
                {
                    this.cellReferences.Add(refToAdd);
                }
            }
        }

        /// <summary>
        /// Removes a reference from the indexer, if it exists.
        /// </summary>
        /// <param name="toRemove">The cell to remove.</param>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="toRemove"/> is null.</exception>
        ////public void Remove(Cell toRemove)
        ////{
        ////    if (toRemove.HasValidCellRef())
        ////    {
        ////        var refToRemove = CellReference.Create(toRemove.CellReference.Value);

        ////        // handle splitting ranges up etcetera.
        ////    }
        ////}

        /// <summary>
        /// Clears all references inside the indexer.
        /// </summary>
        public void Clear()
        {
            this.cellReferences.Clear();
            this.DataValidation.SequenceOfReferences.Items.Clear();
        }

        /// <summary>
        /// Check to see if this indexer contains a reference to the target cell.
        /// </summary>
        /// <param name="toCheck">The cell to check for.</param>
        /// <returns>
        /// True if the validator contains a reference to the target cell, false otherwise.
        /// </returns>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="toCheck"/> is null.</exception>
        public bool Contains(Cell toCheck)
        {
            return toCheck.HasValidCellRef() && this.cellReferences.Contains(CellReference.Create(toCheck));
        }

        #endregion
    }
}
