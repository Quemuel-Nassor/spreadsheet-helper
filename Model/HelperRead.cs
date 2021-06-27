using NPOI.SS.UserModel;
using System;

namespace spreadsheet_helper.Model
{
    /// <summary>
    /// Helper class to read content from spreadsheet cell
    /// </summary>
    public class HelperRead
    {
        /// <summary>
        ///  Get the data from a cell in the line provided
        /// </summary>
        /// <param name="currCell"> Current cell in spreadsheet </param>
        /// <param name="type"> Current cell type </param>
        /// <param name="rowNum"> Current cell number </param>
        /// <param name="headerCell"> Current cell header </param>
        public HelperRead(ICell currentCell, CellType type, int rowNum, string headerCell)
        {
            if (CurrentCell is null)
            {
                throw new ArgumentNullException(nameof(CurrentCell), "Worksheet cell is null or invalid");
            }

            if (HeaderCell is null)
            {
                throw new ArgumentNullException(nameof(Type), "The corresponding spreadsheet header cell is null or invalid");
            }

            CurrentCell = currentCell;
            Type = type;
            RowNum = rowNum;
            HeaderCell = headerCell;
        }

        /// <summary>
        /// Returns cell content as int or null if it is invalid or empty
        /// </summary>
        /// <exception cref="InvalidCastException"> If value conversion fails </exception>
        public int? GetCellAsInt()
        {
            try
            {
                return this.IsNotNull ? (int?)Convert.ToInt32(CurrentCell.ToString()) : null;
            }
            catch
            {
                throw new InvalidCastException("There was an error in the line " + RowNum + " column " + HeaderCell + @" when converting the value '" + CurrentCell.ToString() + @"' to string");
            }
        }

        /// <summary>
        /// Returns cell content as long or null if it is invalid or empty
        /// </summary>
        /// <exception cref="InvalidCastException"> If value conversion fails </exception>
        public long? GetCellAsLong()
        {
            try
            {
                return this.IsNotNull ? (long?)Convert.ToInt64(CurrentCell.ToString()) : null;
            }
            catch
            {
                throw new InvalidCastException("There was an error in the line " + RowNum + " column " + HeaderCell + @" when converting the value '" + CurrentCell.ToString() + @"' to long");
            }
        }

        /// <summary>
        /// Returns cell contents as float or null if it is invalid or empty
        /// </summary>
        /// <exception cref="InvalidCastException"> If value conversion fails </exception>
        public float? GetCellAsFloat()
        {
            try
            {
                return this.IsNotNull ? (float?)float.Parse(CurrentCell.ToString()) : null;
            }
            catch
            {
                throw new InvalidCastException("An error occurred in the line " + RowNum + " column " + HeaderCell + @" when converting the value '" + CurrentCell.ToString() + @"' to float");
            }
        }

        /// <summary>
        /// Returns cell contents as double or null if it is invalid or empty
        /// </summary>
        /// <exception cref="InvalidCastException"> If value conversion fails </exception>
        public double? GetCellAsDouble()
        {
            try
            {
                return this.IsNotNull ? (double?)Convert.ToDouble(CurrentCell.ToString()) : null;
            }
            catch
            {
                throw new InvalidCastException("An error occurred in the line " + RowNum + " column " + HeaderCell + @" when converting the value '" + CurrentCell.ToString() + @"' to double");
            }
        }

        /// <summary>
        /// Returns cell contents as decimal or null if it is invalid or empty
        /// </summary>
        /// <exception cref="InvalidCastException"> If value conversion fails </exception>
        public decimal? GetCellAsDecimal()
        {
            try
            {
                return this.IsNotNull ? (decimal?)Convert.ToDecimal(CurrentCell.ToString()) : null;
            }
            catch
            {
                throw new InvalidCastException("An error occurred in the line " + RowNum + " column " + HeaderCell + @" when converting the value '" + CurrentCell.ToString() + @"' to decimal");
            }
        }

        /// <summary>
        /// Returns cell content as boolean or null if it is invalid or empty
        /// </summary>
        /// <exception cref="InvalidCastException"> If value conversion fails </exception>
        public bool? GetCellAsBool()
        {
            try
            {
                return this.IsNotNull ? (bool?)Convert.ToBoolean(CurrentCell.ToString()) : null;
            }
            catch
            {
                throw new InvalidCastException("There was an error in the line " + RowNum + " column " + HeaderCell + @" when converting the value '" + CurrentCell.ToString() + @"' to Boolean");
            }
        }

        /// <summary>
        /// Returns cell content as a string or null if it is invalid or empty
        /// </summary>
        /// <exception cref="InvalidCastException"> If value conversion fails </exception>
        public string GetCellAsString()
        {
            try
            {
                return this.IsNotNull ? CurrentCell.ToString() : null;
            }
            catch
            {
                throw new InvalidCastException("There was an error in the line " + RowNum + " column " + HeaderCell + @" when converting the value '" + CurrentCell.ToString() + @"' to string");
            }
        }

        public ICell CurrentCell { get; set; }
        public bool IsNotNull { get; set; }

        /// <summary>
        /// True if it is a number represented by a string
        /// </summary>
        public bool IsNumAsString { get; set; }

        /// <summary>
        /// True if it is a number that can be converted to a string
        /// </summary>
        public bool IsStringAsNum { get; set; }
        public bool IsTypeNum { get; set; }
        public bool IsTypeBlank { get; set; }
        public bool IsTypeStr { get; set; }
        public bool IsTypeBool { get; set; }
        public CellType Type { get; set; }
        public int RowNum { get; set; }
        public string HeaderCell { get; set; }

    }
}

