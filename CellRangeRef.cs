using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Aron.Sinoai.OfficeHelper
{
    public struct CellRangeRef
    {
        private static Regex PATTERN = new Regex(@"^(?:(.*?)\!)?(?:\$)?([A-Z]+)(?:\$)?([0-9]+)(?:\:(?:\$)?([A-Z]+)(?:\$)?([0-9]+))?$");

        private string sheetName;
        private CellRef start;
        private CellRef end;

        #region constructors

        public CellRangeRef(string value)
        {
            var match = PATTERN.Match(value);
            sheetName = match.Groups[1].Value;
            start = new CellRef(match.Groups[3].Value, match.Groups[2].Value);

            string endColumn = match.Groups[4].Value;

            if (endColumn.Length == 0)
            {
                end = new CellRef(start);
            }
            else
            {
                end = new CellRef(match.Groups[5].Value, endColumn);
            }
        }

        public CellRangeRef(string sheetName, CellRef start, CellRef end)
        {
            this.sheetName = sheetName;

            this.start = new CellRef(start);
            this.end = new CellRef(end);

            Reorder();
        }

        #endregion

        #region private methods

        private void Reorder()
        {
            if (start.Row > end.Row)
            {
                int temp = start.Row;
                start.Row = end.Row;
                end.Row = temp;
            }

            if (start.Column > end.Column)
            {
                int temp = start.Column;
                start.Column = end.Column;
                end.Column = temp;
            }

        }

        #endregion

        #region public properties

        public String SheetName
        {
            get { return sheetName; }
            set { sheetName = value; }
        }

        public CellRef Start
        {
            get { return start; }
            set { start = value; }
        }

        public CellRef End
        {
            get { return end; }
            set { end = value; }
        }

        public int Height
        {
            get { return end.Row - start.Row; }
        }

        public int Width
        {
            get { return end.Column - start.Column; }
        }

        #endregion

        #region public methods


        override public string ToString()
        {
            string result;

            if (start.Equals(end))
            {
                result = start.ToString();
            }
            else
            {
                result = string.Format("{0}:{1}", start, end);
            }

            if (sheetName.Length > 0)
            {
                result = string.Format("{0}!{1}", sheetName, result);
            }

            return result;
        }

        public bool Contains(CellRef item)
        {
            bool result = start.Row <= item.Row && item.Row <= end.Row &&
                          start.Column <= item.Column && item.Column <= end.Column;

            return result;
        }

        #endregion
    }
}
