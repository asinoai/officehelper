using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.Text.RegularExpressions;

namespace Aron.Sinoai.OfficeHelper
{
    public class ExcelHelper : IDisposable
    {
        
        #region public types

        public enum DirectionType
        {
            LEFT_TO_RIGHT
            ,TOP_TO_DOWN
        }


        #endregion


        #region private memebers

        static private Regex PARAM_PATTERN = new Regex(@"^\<(.*)\>$");  

        private SpreadsheetDocument document;
        private string currentSheetName;
        private SheetData currentSheetData;
        private CellRef currentPosition;
        private Row currentRow; //actually it is the row before which we should insert the current-row
        private DirectionType direction;
  

        #endregion

        #region contructors

        public ExcelHelper(string templateFileName, string generatedFileName)
        {
            File.Copy(templateFileName, generatedFileName, true /*overwrite*/);

            document = SpreadsheetDocument.Open(generatedFileName, true/*isEditable*/);

            currentPosition = new CellRef(1, 1);//1 based indexing

            direction = DirectionType.TOP_TO_DOWN;
        }

        #endregion

        #region public methods

        public void Dispose()
        {
            document.Dispose();
            document = null;
        }
        
        public void InsertRange(CellRangeTemplate template)
        {
            InsertRange(template.Range);
        }

        public void InsertRange(CellRangeTemplate template, List<object> entity)
        {
            template.Assign(entity);
            InsertRange(template.Range);
        }

        public void InsertRange(CellRangeTemplate template, IEnumerable<List<object>> entities)
        {
            foreach (var entity in entities)
            {
                template.Assign(entity);
                InsertRange(template.Range);
            }
        }

        public void InsertRange(string name)
        {
            CellRangeRef range = FindDefinedNameRange(name);
            InsertRange(range);
        }

        public void InsertRange(CellRangeRef range)
        {
            CopyRange(ref range, currentSheetData, ref currentPosition);

            if (direction == DirectionType.TOP_TO_DOWN)
            {
                currentPosition.Row = currentPosition.Row + range.Height + 1;
            }
            else
            {
                currentPosition.Column = currentPosition.Column + range.Width + 1;
            }
        }

        public void SetCurrentPositionByName(string definedName)
        {
            CellRangeRef range = FindDefinedNameRange(definedName);
            CurrentPosition = range.Start;
        }

        public CellRangeRef FindDefinedNameRange(string name)
        {
            DefinedName definedName = (
                from item in document.WorkbookPart.Workbook.DefinedNames.Elements<DefinedName>()
                where item.Name == name
                select item).Single();
            CellRangeRef range = new CellRangeRef(definedName.Text);
            return range;
        }

        public void DeleteSheet(string name)
        {
            FindSheetByName(name).Remove();
        }

        public CellRangeTemplate CreateCellRangeTemplate(string name)
        {
            CellRangeRef range = FindDefinedNameRange(name);
            return CreateCellRangeTemplate(range);
        }

        public CellRangeTemplate CreateCellRangeTemplate(string name, List<string> names)
        {
            CellRangeTemplate result = CreateCellRangeTemplate(name);
            result.BuildNumericIndexing(names);

            return result;
        }

        public CellRangeTemplate CreateCellRangeTemplate(CellRangeRef range)
        {
            CellRangeTemplate result = new CellRangeTemplate(range);

            var cells = FindCellsByRange(range);
            foreach (var rowGroup in cells)
            {
                foreach (Cell cell in rowGroup)
                {
                    if (cell.DataType != null && cell.CellValue != null &&
                        cell.DataType.Value == CellValues.SharedString)
                    {
                        int index = int.Parse(cell.CellValue.Text);
                        SharedStringItem stringItem = document.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(index);

                        string stringValue = stringItem.InnerText;
                        if (PARAM_PATTERN.IsMatch(stringValue))
                        {
                            var match = PARAM_PATTERN.Match(stringValue);
                            var paramName = match.Groups[1].Value;
                            result.AddCell(paramName, cell);
                        }
                    }
                }
            }

            return result;
        }

        #endregion

        #region public properties

        public string CurrentSheetName
        {
            get { return currentSheetName; }
            set 
            { 
                currentSheetName = value;
                currentSheetData = FindSheetDataByName(currentSheetName);
            }
        }


        public CellRef CurrentPosition
        {
            get { return currentPosition; }
            set 
            { 
                currentPosition = value;

                //reseting this since this is a cached member
                currentRow = null;
            }
        }

        public DirectionType Direction
        {
            get { return direction; }
            set { direction = value; }
        }

        #endregion

        #region private methods

        private IEnumerable<IGrouping<Row, Cell>> FindCellsByRange(CellRangeRef range)
        {
            Sheet sheet = FindSheetByName(range.SheetName);

            WorksheetPart workSheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id);

            SheetData sheetData = workSheetPart.Worksheet.GetFirstChild<SheetData>();

            var result = from row in sheetData.Elements<Row>()
                         where range.Start.Row <= row.RowIndex && row.RowIndex <= range.End.Row
                         let cells = row.Elements<Cell>()
                         from cell in cells
                         where range.Contains(new CellRef(cell.CellReference))
                         group cell by row into cellsInRange
                         select cellsInRange;

            return result;
        }

        private Sheet FindSheetByName(string sheetName)
        {
            Sheet sheet =
                (from item in document.WorkbookPart.Workbook.Sheets.Elements<Sheet>()
                 where item.Name == sheetName
                 select item).Single();
            return sheet;
        }

        private SheetData FindSheetDataByName(string sheetName)
        {
            Sheet sheet = FindSheetByName(sheetName);

            SheetData sheetData = FindSheetDataBySheet(sheet);

            return sheetData;
        }

        private SheetData FindSheetDataBySheet(Sheet sheet)
        {
            WorksheetPart workSheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id);

            SheetData sheetData = workSheetPart.Worksheet.GetFirstChild<SheetData>();
            return sheetData;
        }

        private void MoveCurrentRow(int rowIndex)
        {
            if (currentRow == null)
            {
                var rows = currentSheetData.Elements<Row>();
                if (rows.GetEnumerator().MoveNext())
                {
                    //maybe the last row is still before the curretRow, so in this case no need to iterate trough the rows
                    Row lastRow = (Row)currentSheetData.LastChild;
                    if (lastRow.RowIndex > rowIndex)
                    {
                        //we search from the beginnig
                        foreach (Row item in rows)
                        {
                            if (item.RowIndex.Value > rowIndex)
                            {
                                currentRow = item;
                                break;
                            }
                        }
                    }
                }
            }
            else
            {
                //trying to seach from the last position backward or forward
                if (currentRow.RowIndex > rowIndex)
                {
                    //going backward
                    Row previousRow;
                    while ((previousRow = currentRow.PreviousSibling<Row>()) != null &&
                            previousRow.RowIndex > rowIndex)
                    {
                        currentRow = previousRow;
                    }

                }
                else
                {
                    //going forward
                    while (currentRow != null && currentRow.RowIndex <= rowIndex)
                    {
                        currentRow = currentRow.NextSibling<Row>();
                    }
                }
            }

        }

        private void CopyRange(ref CellRangeRef sourceRange, SheetData sheetData, ref CellRef target)
        {
            CellRef source = sourceRange.Start;
            CellRef offset = target.CalculateOffset(source);

            var cellsToCopy = FindCellsByRange(sourceRange);
            foreach (var rowGroup in cellsToCopy)
            {
                Row keyRow = rowGroup.Key;

                Row targetRow = new Row()
                {
                    RowIndex = (UInt32)(keyRow.RowIndex + offset.Row)
                };

                MoveCurrentRow((int)targetRow.RowIndex.Value);
                sheetData.InsertBefore(targetRow, currentRow);

                foreach (Cell cellToCopy in rowGroup)
                {
                    Cell targetCell = (Cell)cellToCopy.Clone();

                    targetCell.CellReference = CellRef.OffsetIt(targetCell.CellReference, offset);

                    targetRow.Append(targetCell);
                }

            }
        }


        #endregion

    }
}
