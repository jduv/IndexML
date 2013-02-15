namespace IndexML.Spreadsheet
{
    using System;
    using System.Text.RegularExpressions;
    using DocumentFormat.OpenXml.Spreadsheet;
    using IndexML.Extensions;

    /// <summary>
    /// Abstract class containing utility methods and base implementations
    /// for cell references.
    /// </summary>
    public abstract class CellReference : ICellReference
    {
        #region Fields & Constants

        /// <summary>
        /// A regex pattern for matching cell references.
        /// </summary>
        public static readonly string SingleCellRefRegexStringStrict = @"^(?<col>[a-zA-Z]{1,3})(?<row>[0-9]+)$";

        /// <summary>
        /// A regex pattern for matching cell references with or without the row number.
        /// </summary>
        public static readonly string SingleCellRefRegexString = @"^(?<col>[a-zA-Z]{1,3})(?<row>[0-9]*)$";

        /// <summary>
        /// A regex pattern for matching range cell references.
        /// </summary>
        public static readonly string RangeCellRefRegexString = @"^(?<s>[a-zA-Z]{1,3}[0-9]+):(?<e>[a-zA-Z]{1,3}[0-9]+)$";

        /// <summary>
        /// A cached regex for matching/replacing single cell references with or without the row index.
        /// </summary>
        public static readonly Regex SingleCellRefRegex = new Regex(SingleCellRefRegexString, RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>
        /// A cached regex for matching/replacing single cell references.
        /// </summary>
        public static readonly Regex SingleCellRefRegexStrict = new Regex(SingleCellRefRegexStringStrict, RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>
        /// A cached regex for matching/replacing range cell references.
        /// </summary>
        public static readonly Regex RangeCellRefRegex = new Regex(RangeCellRefRegexString, RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>
        /// The alphabet.
        /// </summary>
        private static readonly char[] Alphabet = new char[] { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' };

        /// <summary>
        /// The cell reference.
        /// </summary>
        private string reference;

        #endregion

        #region Constructors & Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="CellReference"/> class. This is a default instance
        /// with no extra initialization performed on the values contained inside--it is the subclasses
        /// responsibility to ensure that the Value property is correctly set.
        /// </summary>
        protected CellReference()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CellReference"/> class.
        /// </summary>
        /// <param name="cellRef">The cell to initialize with.</param>
        /// <exception cref="ArgumentException">Thrown if <paramref name="cellRef"/> is null, empty,
        /// or invalid.</exception>
        protected CellReference(string cellRef)
        {
            if (string.IsNullOrEmpty(cellRef))
            {
                throw new ArgumentException("Invalid cell reference:" + cellRef.PrettyPrint());
            }

            if (IsValidCellReference(cellRef))
            {
                this.reference = cellRef;
            }
            else
            {
                throw new ArgumentException("Invalid cell reference: " + cellRef);
            }
        }

        #endregion

        #region Properties

        /// <inheritdoc/>
        public string Value
        {
            get
            {
                if (string.IsNullOrEmpty(this.reference))
                {
                    throw new InvalidOperationException("The value for this reference hasn't been set properly!");
                }

                return this.reference;
            }

            protected set
            {
                if (string.IsNullOrEmpty(value) || !IsValidCellReference(value))
                {
                    throw new InvalidOperationException("The value for this reference is invalid: " + value.PrettyPrint());
                }

                this.reference = value;
            }
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Determines if the passed in string is a valid cell reference.
        /// </summary>
        /// <param name="toExamine">The string to examine.</param>
        /// <returns>True if the string is a valid cell reference, false otherwise.</returns>
        public static bool IsValidCellReference(string toExamine)
        {
            return IsSingleCellReference(toExamine) || IsRangeCellReference(toExamine);
        }

        /// <summary>
        /// Determines if the passed in string is a valid single cell reference as dicated by the
        /// regular expression for single cell references.
        /// </summary>
        /// <param name="toExamine">The string to examine.</param>
        /// <returns>True if the string is a valid single cell reference, false otherwise.</returns>
        public static bool IsSingleCellReference(string toExamine)
        {
            if (string.IsNullOrEmpty(toExamine))
            {
                return false;
            }
            else
            {
                return SingleCellRefRegexStrict.Match(toExamine).Success;
            }
        }

        /// <summary>
        /// Determines if the passed in string is a valid range cell reference as dicated by the
        /// regular expression for range cell references.
        /// </summary>
        /// <param name="toExamine">The string to examine.</param>
        /// <returns>True if the string is a valid range cell reference, false otherwise.</returns>
        public static bool IsRangeCellReference(string toExamine)
        {
            if (string.IsNullOrEmpty(toExamine))
            {
                return false;
            }
            else
            {
                return RangeCellRefRegex.Match(toExamine).Success;
            }
        }

        /// <summary>
        /// Creates a cell reference from the target cell.
        /// </summary>
        /// <param name="cell">The cell to create the reference for.</param>
        /// <returns>Returns the proper cell reference implementation for the given cell, or throws an 
        /// exception if one doesn't exist.</returns>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="cell"/> is null.</exception>
        /// <exception cref="ArgumentException">Thrown if the cell has an invalid cell reference.</exception>
        public static ICellReference Create(Cell cell)
        {
            if (cell == null)
            {
                throw new ArgumentNullException("cell");
            }

            if (cell.CellReference == null)
            {
                throw new ArgumentException("Invalid cell reference! It's null!");
            }

            return Create(cell.CellReference.Value);
        }

        /// <summary>
        /// Creates a cell reference from the target string.
        /// </summary>
        /// <param name="cellRef">The string to parse.</param>
        /// <returns>A cell reference implementation for the given reference string.</returns>
        /// <exception cref="ArgumentException">Thrown if the target <paramref name="cellRef"/> is invalid.</exception>
        public static ICellReference Create(string cellRef)
        {
            if (string.IsNullOrEmpty(cellRef))
            {
                throw new ArgumentException("Invalid cell reference: (" + (cellRef == null ? "null" : "empty") + ").");
            }
            else if (IsSingleCellReference(cellRef))
            {
                return new SingleCellReference(cellRef);
            }
            else if (IsRangeCellReference(cellRef))
            {
                return new RangeCellReference(cellRef);
            }
            else
            {
                throw new ArgumentException("Invalid cell reference: " + cellRef);
            }
        }

        /// <summary>
        /// Parses the column reference from the given arbitrary string and returns a column index.
        /// </summary>
        /// <param name="columnName">The string to parse the column index from.</param>
        /// <param name="strict">Indicates whether to perform a strict match on a cell reference type or not.
        /// Dictates which regular expression to use when matching cell reference patterns.</param>
        /// <param name="colIdx">The column index to assign to.</param>
        /// <returns>True if the column index for the target string was parsed successfully, false otherwise..</returns>
        public static bool TryGetColumnIndex(string columnName, bool strict, out long colIdx)
        {
            if (!string.IsNullOrEmpty(columnName))
            {
                string parsedColName;
                if (TryGetColumnName(columnName, strict, out parsedColName))
                {
                    int number = 0;
                    int pow = 1;
                    for (int i = parsedColName.Length - 1; i >= 0; i--)
                    {
                        number += (parsedColName[i] - 'A' + 1) * pow;
                        pow *= 26;
                    }

                    colIdx = number;
                    return true;
                }
            }

            colIdx = default(int);
            return false;
        }

        /// <summary>
        /// Given an input string, attempt to parse it into a column name based on a regular expression.
        /// </summary>
        /// <param name="inputString">The input string.</param> 
        /// <param name="strict">Indicates whether to perform a strict match on a cell reference type or not.
        /// Dictates which regular expression to use when matching cell reference patterns.</param>
        /// <param name="colName">The column name to assign to.</param>
        /// <returns>True if the assignment was successful, false othewise.</returns>
        public static bool TryGetColumnName(string inputString, bool strict, out string colName)
        {
            if (!string.IsNullOrEmpty(inputString))
            {
                var match = strict ? SingleCellRefRegexStrict.Match(inputString) : SingleCellRefRegex.Match(inputString);
                if (match.Success)
                {
                    colName = match.Groups["col"].Value;
                    return true;
                }
            }

            colName = default(string);
            return false;
        }

        /// <summary>
        /// Converts the target column index into the approriate Excel column name.
        /// </summary>
        /// <param name="colIdx">The column index to convert.</param>
        /// <returns>A string corresponding to the correct column index.</returns>
        public static string GetColumnName(long colIdx)
        {
            if (colIdx <= 0)
            {
                throw new ArgumentOutOfRangeException(
                    "Invalid column index. " + colIdx + ". Remember that column indices are one based.");
            }

            var dividend = colIdx;
            var columnName = string.Empty;
            long modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        /// <summary>
        /// Given an input string, attempt to parse it into a row index based on a regular expression.
        /// </summary>
        /// <param name="inputString">The input string.</param>
        /// <param name="rowIdx">The row index to assign to.</param>
        /// <returns>The row index.</returns>
        public static bool TryGetRowIndex(string inputString, out long rowIdx)
        {
            // BMK: Add strict bool flag to this method.
            if (!string.IsNullOrEmpty(inputString))
            {                
                var match = SingleCellRefRegexStrict.Match(inputString);
                if (match.Success && long.TryParse(match.Groups["row"].Value, out rowIdx))
                {
                    return true;
                }
            }

            rowIdx = default(int);
            return false;
        }

        /// <inheritdoc/>
        public abstract bool ContainsOrSubsumes(ICellReference cellRef);
        
        /// <inheritdoc />
        public abstract ICellReference Resize(int rows, int cols);                

        /// <inheritdoc />
        public abstract ICellReference Move(int rows, int cols);        

        /// <summary>
        /// Checks to see if the value contained in one cell reference is exactly equal to the value contained in
        /// the other cell reference. Use this for quick comparisons or testing, otherwise use Equals, which calls
        /// this method anyhow.
        /// </summary>
        /// <param name="first">The first cell reference to compare.</param>
        /// <param name="second">The second cell reference to compare.</param>
        /// <returns>True if the value properties of this and the other cell reference are exactly equal,
        /// false otherwise.</returns>
        public static bool ValueEquals(ICellReference first, ICellReference second)
        {
            // We should only compare valid cell references. Nulls mean no equality.
            if (first == null || second == null)
            {
                return false;
            }

            return first.Value.Equals(second.Value, StringComparison.OrdinalIgnoreCase);
        }

        /// <inheritdoc />
        public override string ToString()
        {
            return this.Value.PrettyPrint();
        }

        #endregion
    }
}
