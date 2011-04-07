using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace Aron.Sinoai.OfficeHelper
{
    public class CellRangeTemplate
    {
        #region public types

        public delegate string CoversionDelegate(object item);

        #endregion


        #region private members

        private IDictionary<string, Cell> dictionay = new Dictionary<string, Cell>();
        private List<Cell> array = null;
        
        private CellRangeRef range;

        #endregion

        #region public interface

        public CellRangeTemplate(CellRangeRef range)
        {
            this.range = range;
        }

        public void AddCell(string name, Cell cell)
        {
            dictionay[name] = cell;
        }

        public string this[string name]
        {
            get
            {
                return dictionay[name].CellValue.Text;
            }

            set
            {
                Cell cell = dictionay[name];
                SetCellValue(value, cell);
            }
        }

        public object this[int index]
        {
            get
            {
                return array[index].CellValue.Text;
            }

            set
            {
                Cell cell = array[index];
                SetCellValue(value, cell);

            }
        }

 
        public CellRangeTemplate Assign(IList<object> entity)
        {
            for (int i = 0; i < entity.Count; i++)
            {
                this[i] = entity[i];
            }

            return this;
        }

        public void BuildNumericIndexing(List<string> names)
        {
            array = new List<Cell>(names.Count);

            for (int i = 0; i < names.Count; i++)
            {
                array.Add(dictionay[names[i]]);
            }
        }

        public CellRangeRef Range
        {
            get
            {
                return range;
            }
        }

        #endregion


        #region private methods

        private static void SetCellValue(object value, Cell cell)
        {
            if (value is Double)
            {
                cell.DataType.Value = CellValues.Number;
                cell.CellValue.Text = ((Double)value).ToString(CultureInfo.InvariantCulture);
            }
            else if (value is Int16)
            {
                cell.DataType.Value = CellValues.Number;
                cell.CellValue.Text = ((Int16)value).ToString(CultureInfo.InvariantCulture);
            }
            else if (value is Int32)
            {
                cell.DataType.Value = CellValues.Number;
                cell.CellValue.Text = ((Int32)value).ToString(CultureInfo.InvariantCulture);
            }
            else if (value is Int64)
            {
                cell.DataType.Value = CellValues.Number;
                cell.CellValue.Text = ((Int64)value).ToString(CultureInfo.InvariantCulture);
            }
            else if (value is DateTime)
            {
                cell.DataType.Value = CellValues.Number;
                cell.CellValue.Text = ((DateTime)value).ToOADate().ToString(CultureInfo.InvariantCulture);
            }
            else
            {
                cell.DataType.Value = CellValues.String;
                cell.CellValue.Text = value.ToString();
            }
        }

        #endregion
    }
}
