using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Xml.Linq;
using System.Data.SQLite;
using Aron.Sinoai.OfficeHelper;

namespace WindowsFormsApplication3
{
    public partial class mainForm : Form
    {
        public mainForm()
        {
            InitializeComponent();
        }

        private void goButton_Click(object sender, EventArgs e)
        {
            export(tableNameComboBox.Text);
        }

        private void dbBrowseButton_Click(object sender, EventArgs e)
        {
            openFileDialog.FileName = dbFileTextBox.Text;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                dbFileTextBox.Text = openFileDialog.FileName;
            }

        }

        private void dbFileTextBox_TextChanged(object sender, EventArgs e)
        {
            if (File.Exists(dbFileTextBox.Text))
            {
                refreshTableNames();
            }
        }

        private SQLiteConnection createConnection()
        {
            SQLiteConnection connection = new SQLiteConnection(string.Format("Data Source={0};FailIfMissing=true", dbFileTextBox.Text));
            connection.Open();
            return connection;
        }

        private List<string> fetchTables(SQLiteConnection connection)
        {
            DataTable columns = connection.GetSchema("Tables");

            IEnumerable<string> result =
                from row in columns.Select()
                select row.Field<String>("TABLE_NAME");

            return result.ToList<string>();
        }

        private List<string> fetchColumns(SQLiteConnection connection, string tableName)
        {
            DataTable columns = connection.GetSchema("Columns");

            IEnumerable<string> result =
                from row in columns.Select(String.Format("TABLE_NAME = '{0}'", tableName))
                select row.Field<String>("COLUMN_NAME");

            return result.ToList<string>();
        }

        private IEnumerable<List<object>> fetchData(SQLiteConnection connection, string tableName, List<string> columns)
        {
            string separator = "";
            StringBuilder buffer = new StringBuilder();
            foreach(string column in columns)
            {
                buffer.Append(separator);
                buffer.Append(column);

                if (separator.Length == 0) 
                {
                    separator = ", ";
                }
            }

            int count = columns.Count;
            List<object> result = new List<object>(count);
            
            SQLiteCommand command = connection.CreateCommand();
            command.CommandText = string.Format("select {0} from {1}", buffer.ToString(), tableName);
            using (SQLiteDataReader reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    result.Clear();
                    for(int i = 0; i < count; i++)
                    {
                        object value = reader.GetValue(i);
                        if (value is string)
                        {
                            string stringValue = (string)value;
                            if (stringValue.Contains((char)0x1a))
                            {
                                stringValue = stringValue.Replace((char)0x1a, ' ');
                                value = stringValue;
                            }
                            else
                            {
                                value = stringValue;
                            }
                        }
                             
                        result.Add(value);


                    }
                    
                    yield return result;
                }
            }
        }

        /*private void init()
        {
            using (TextReader patchContentReader = new StreamReader(@"resources\database.xml") )
            {
                XElement root = XElement.Load(patchContentReader);

                var result =
                    from item in root.Element("tables").Elements("item")
                    select item.Attribute("name").Value;

                tableNameComboBox.DataSource = result.ToList<String>();
            }
        }*/

        private void refreshTableNames()
        {
            SQLiteConnection connection = createConnection();
            using (connection)
            {
                List<string> tables = fetchTables(connection);
                tables.Sort();
                tableNameComboBox.DataSource = tables;

            }
        }

        private void export(string tableName)
        {
            using (SQLiteConnection connection = createConnection())
            {
                List<string> columns = fetchColumns(connection, tableName);

                string appFolder = Path.GetDirectoryName(Application.ExecutablePath);
                string templateFileName = Path.Combine(appFolder, string.Format(@"resources\{0}.xlsx", tableName));
                string exportFileName = string.Format("{0}.xlsx", tableName);

                if (File.Exists(templateFileName))
                {
                    using (ExcelHelper helper = new ExcelHelper(templateFileName, exportFileName))
                    {
                        helper.Direction = ExcelHelper.DirectionType.TOP_TO_DOWN;
                        helper.CurrentSheetName = "Sheet1";

                        helper.InsertRange("header");
                        CellRangeTemplate rowTemplate = helper.CreateCellRangeTemplate("row", columns);

                        helper.InsertRange(rowTemplate, fetchData(connection, tableName, columns));

                        helper.DeleteSheet("Templates");
                    }

                    MessageBox.Show("Exported!");
                }
                else
                {
                    MessageBox.Show(String.Format("Template xlsx not found ('{0}')!\nAutocreating it, please check it and retry!", templateFileName));
                    string basicTempateFileName = Path.Combine(appFolder, @"resources\BasicTemplate.xlsx");
                    using (ExcelHelper helper = new ExcelHelper(basicTempateFileName, templateFileName))
                    {
                        helper.Direction = ExcelHelper.DirectionType.LEFT_TO_RIGHT;
                        helper.CurrentSheetName = "Templates";

                        CellRangeTemplate template = helper.CreateCellRangeTemplate("header_row", new List<string>() { "header", "row" });

                        var result = 
                            from item in columns
                            select new List<object>() { item, string.Format("<{0}>", item) };

                        helper.InsertRange(template, result);

                        var names = new string[] {"header", "row"};
                        foreach (string name in names)
                        {
                            CellRangeRef range = helper.FindDefinedNameRange(name);
                            
                            //extending it to the length of the columns
                            CellRef end = range.End;
                            end.OffsetIt(new CellRef(0, columns.Count - 1));
                            range.End = end;
                            range.SheetName = "Templates";

                            helper.SetDefinedNameRange(name, range);
                        }

                        helper.DeleteSheet("_Templates");
                    }
                }

            }
        }

    }
}
