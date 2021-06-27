using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using spreadsheet_helper.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace spreadsheet_helper.Manager
{
    public class SpreadsheetManager
    {
        /// <summary>
        /// Method to export content to spreasheet file
        /// </summary>
        /// <param name="content">List of content to be inserted</param>
        /// <param name="sheetName">spreadsheet name</param>
        /// <returns></returns>
        public void ExportSpreadsheet<T>(List<T> content, string sheetName, List<HelperColumns> columns)
        {
            try
            {
                #region Spreadsheet Start
                XSSFWorkbook workbook = new XSSFWorkbook();
                ICellStyle headerStyle = workbook.CreateCellStyle();
                ICellStyle bodyStyle = workbook.CreateCellStyle();
                IFont fontBody = workbook.CreateFont();
                IFont fontHeader = workbook.CreateFont();

                //Setting font style
                fontBody.FontName = XSSFFont.DEFAULT_FONT_NAME;
                fontHeader.FontName = XSSFFont.DEFAULT_FONT_NAME;
                fontHeader.IsBold = true;
                fontHeader.Color = HSSFColor.White.Index;

                //Setting body style
                bodyStyle.BorderBottom = BorderStyle.Thin;
                bodyStyle.BorderTop = BorderStyle.Thin;
                bodyStyle.BorderLeft = BorderStyle.Thin;
                bodyStyle.BorderRight = BorderStyle.Thin;
                bodyStyle.SetFont(fontBody);
                bodyStyle.Alignment = HorizontalAlignment.Center;
                bodyStyle.VerticalAlignment = VerticalAlignment.Center;

                //Setting header style
                headerStyle.CloneStyleFrom(bodyStyle);
                headerStyle.FillForegroundColor = HSSFColor.DarkRed.Index2;
                headerStyle.FillPattern = FillPattern.SolidForeground;
                headerStyle.SetFont(fontHeader);
                #endregion

                const int HeaderRow = -1;

                if (columns == null || columns.Count() == 0)
                {
                    foreach (var p in typeof(T).GetProperties())
                    {
                        columns.Add(
                            new HelperColumns(p, bodyStyle, p.Name, true)
                        );
                    }
                }

                ISheet Sheet = workbook.CreateSheet(sheetName);

                //Inserting data
                for (int i = HeaderRow; i < content.Count; i++)
                {
                    IRow Row = Sheet.CreateRow(i + 1);

                    for (int j = 0; j < columns.Count(); j++)
                    {
                        var c = columns[j];
                        _ = i == HeaderRow ?
                        new HelperWrite(ref Row, headerStyle, j, c.Title, null) :
                        new HelperWrite(ref Row, c.Style, j, c.ClassProperty.GetValue(content[i]), null);
                    }
                }

                //var result = new FileDto("Exported Content.xlsx", MimeTypeNames.ApplicationVndOpenxmlformatsOfficedocumentSpreadsheetmlSheet);
                //Save(workbook, result);

            }
            catch
            {
                throw new Exception("An error occurred while exporting the contents of the informed list");
            }
        }

        /// <summary>
        /// Method to load content from spreadsheet file
        /// </summary>
        /// <typeparam name="T"> Type of content of destinatation </typeparam>
        /// <param name="fileSourcePath"> File path </param>
        /// <returns> Collection of lists where each list contains content of sheet </returns>
        public List<List<T>> ImportSpreadsheet<T>(string fileSourcePath)
        {
            try
            {
                List<List<T>> contentsCollection = new List<List<T>>();
                List<T> contentList = new List<T>();

                XSSFWorkbook workbook;

                //Load spreadsheet stream
                using (var doc = File.OpenRead(fileSourcePath))
                {
                    workbook = new XSSFWorkbook(doc);
                }

                //Walk through spreadsheet sheets
                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {
                    var sheet = workbook.GetSheetAt(i);

                    //Walk through rows of sheet
                    for (int rowNum = sheet.FirstRowNum + 1; rowNum < sheet.PhysicalNumberOfRows; rowNum++)
                    {
                        var row = sheet.GetRow(rowNum);

                        T item = (T)new object();
                        
                        if (row == null || row.LastCellNum == -1) throw new ArgumentNullException(nameof(row), "Invalid or empty row on spreadsheet");

                        //Walk through columns of current row
                        for (var colNum = row.FirstCellNum; colNum < row.PhysicalNumberOfCells; colNum++)
                        {
                            string header = sheet.GetRow(sheet.FirstRowNum).GetCell(colNum).StringCellValue;
                            string throwMessage = "Valor incorreto na linha " + (rowNum + 1) + @" coluna '" + header + @"'";
                            ICell currentCell = sheet.GetRow(rowNum).GetCell(colNum);

                            HelperRead cell = new HelperRead(currentCell, currentCell.CellType, rowNum + 1, header);

                            var type = item.GetType().GetProperty(header).PropertyType;

                            if (type == typeof(int)) item.GetType().GetProperty(header).SetValue(item, cell.GetCellAsInt());
                            if (type == typeof(long)) item.GetType().GetProperty(header).SetValue(item, cell.GetCellAsLong());
                            if (type == typeof(double)) item.GetType().GetProperty(header).SetValue(item, cell.GetCellAsDouble());
                            if (type == typeof(float)) item.GetType().GetProperty(header).SetValue(item, cell.GetCellAsFloat());
                            if (type == typeof(decimal)) item.GetType().GetProperty(header).SetValue(item, cell.GetCellAsDecimal());
                            if (type == typeof(bool)) item.GetType().GetProperty(header).SetValue(item, cell.GetCellAsBool());
                            if (type == typeof(string)) item.GetType().GetProperty(header).SetValue(item, cell.GetCellAsString());

                        }

                        contentList.Add(item);

                    }

                    contentsCollection.Add(contentList);
                }

                return contentsCollection;
            }
            catch (Exception error)
            {
                if (error.GetType().Name != "Exception")
                    throw error;
                throw new Exception("Ocorreu um erro ao carregar o arquivo para importação");
            }
        }
    }
}
