using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;

namespace spreadsheet_helper.Model
{
    /// <summary>
    /// Helper class to write content into spreadsheet cell
    /// </summary>
    public class HelperWrite
    {
        /// <summary>
        /// Adds a cell with formatting and data in the given row
        /// </summary>
        /// <param name="currRow"> Current worksheet row </param>
        /// <param name="style"> Cell style </param>
        /// <param name="colNum"> Column number </param>
        /// <param name="value"> Content to be inserted </param>
        /// <param name="dataFormat"> Formatting for the inserted data </param>
        /// <exception cref="ArgumentNullException"> If one or more required parameters are null or invalid </exception>
        public HelperWrite(ref IRow currRow, ICellStyle style, int colNum, dynamic value, int? dataFormat)
        {
            if (currRow is null)
            {
                throw new ArgumentNullException(nameof(currRow),"Worksheet reference line is null or invalid");
            }

            if (style is null)
            {
                throw new ArgumentNullException(nameof(style),"Worksheet style is null or invalid");
            }

            if (value is null)
            {
                throw new ArgumentNullException(nameof(value),"Content to write into spreadsheet cell is null or invalid");
            }

            //Copying cell style
            ICellStyle clonedStyle = new XSSFWorkbook().CreateCellStyle();
            clonedStyle.CloneStyleFrom(style);

            //Creating stylized cell
            this.CurrCell = currRow.CreateCell(colNum);
            this.CurrCell.CellStyle = clonedStyle;

            //Validating and setting the cell value
            if (value != null) this.SetCellType(value.GetType());
            this.SetCellValue(value);

            //Defining the value representation format
            if (dataFormat != null) SetDataFormat(dataFormat);
            currRow.Sheet.AutoSizeColumn(colNum);
        }

        /// <summary>
        /// Sets the value in the cell if it is provided and different from null
        /// </summary>
        /// <param name="value"> Content to be inserted </param>
        private void SetCellValue(dynamic value)
        {
            if (value != null)
            {
                var type = value.GetType();
                if (type.Equals(typeof(int)) || type.Equals(typeof(decimal)) || type.Equals(typeof(long)) || type.Equals(typeof(float)) || type.Equals(typeof(double)))
                {
                    this.CurrCell.SetCellValue((double)value);
                }
                else
                {
                    this.CurrCell.SetCellValue(value);
                }
            }

        }

        /// <summary>
        /// Defines formatting for the cell
        /// </summary>
        /// <param name="format">1 for%, 2 for $</param>
        private void SetDataFormat(int? format)
        {
            switch (format)
            {
                case 1:
                    {
                        //Percentage format
                        this.CurrCell.CellStyle.DataFormat = 9;
                        break;
                    }
                case 2:
                    {
                        //Monetary format
                        this.CurrCell.CellStyle.DataFormat = 7;
                        break;
                    }
                default:
                    break;
            }
        }

        /// <summary>
        /// Defines the data type of the cell based on the type of data being inserted
        /// </summary>
        /// <param name="type"> Data type </param>
        private void SetCellType(Type type)
        {
            if (type.Equals(typeof(int)) || type.Equals(typeof(decimal)) || type.Equals(typeof(long)) || type.Equals(typeof(float)) || type.Equals(typeof(double)))
            {
                this.CurrCell.SetCellType(CellType.Numeric);
            }
            else if (type.Equals(typeof(bool)))
            {
                this.CurrCell.SetCellType(CellType.Boolean);
            }
            else if (type.Equals(typeof(string)))
            {
                this.CurrCell.SetCellType(CellType.String);
            }
            else
            {
                this.CurrCell.SetCellType(CellType.Blank);
            }

        }

        public ICell CurrCell { get; set; }
    }
}
