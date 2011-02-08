using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Aron.Sinoai.OfficeHelper
{
    public struct CellRef
    {
        private static Regex PATTERN = new Regex(@"^([A-Z]+)([0-9]+)$");

        private int row;
        private int column;

        #region constructors

        public CellRef(string value)
        {
            var match = PATTERN.Match(value);
            column = Column2Int(match.Groups[1].Value);
            row = int.Parse(match.Groups[2].Value);
        }

        public CellRef(int row, int column)
        {
            this.row = row;
            this.column = column;
        }

        public CellRef(string row, string column)
        {
            this.row = int.Parse(row);
            this.column = Column2Int(column);
        }

        public CellRef(CellRef right)
        {
            row = right.row;
            column = right.column;
        }

        #endregion

        #region static methods

        static public int Column2Int(string value)
        {
            int result = 0;

            if (value != null && value.Length > 0)
            {
                foreach (char item in value)
                {
                    result = 26 * result + (item - 'A');
                }

                result++; //1 based

                if (value.Length > 1)
                {
                    result += 26;
                }
            }

            return result;
        }

        static public string Int2Column(int value)
        {
            List<char> result = new List<char>();

            if (value > 0)
            {
                value--; //1 based

                if (value < 26)
                {
                    result.Add((char)('A' + value));
                }
                else
                {
                    value -= 26;
                    do
                    {
                        result.Add((char)('A' + value % 26));
                        value = value / 26;
                    } while (value > 0);

                    if (result.Count == 1)
                    {
                        result.Add('A');
                    }
                }
            }

            return new string(result.Reverse<char>().ToArray());
        }


        static public string ToString(int row, int column)
        {
            return string.Format("{0}{1}", CellRef.Int2Column(column), row);
        }

        static public string OffsetIt(string cellRef, CellRef offset)
        {
            CellRef result = new CellRef(cellRef);
            result.OffsetIt(offset);
            return result.ToString();
        }

        #endregion

        #region public interface

        public override string ToString()
        {
            return ToString(row, column);
        }

        public override bool Equals(object obj)
        {
            CellRef right = (CellRef)obj;
            return (row == right.row && column == right.Column);
        }

        public override int GetHashCode()
        {
            return row.GetHashCode() ^ column.GetHashCode();
        }

        public int Row
        {
            get { return row; }
            set { row = value; }
        }

        public int Column
        {
            get { return column; }
            set { column = value; }
        }

        public CellRef CalculateOffset(CellRef value)
        {
            return new CellRef(row - value.row, column - value.column);
        }

        public void OffsetIt(CellRef value)
        {
            row += value.row;
            column += value.column;
        }

        #endregion

    }
}
