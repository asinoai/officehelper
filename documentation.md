You can check the SqliteExporter folder in the Source Code section, for a more complex example, which uses in action this library, or you can check the below sample.
The SqliteExporter is even available as executable in the download section.

Sample code to use the library

What is great by using this library is that you don't need to have Excel (or Microsoft Office) installed on your computer.
Of course to create your template xlsx files you need Excel, but you can deploy it and use it without the need to have Exel on the target deployment machine.
The reason for this is that the library uses the open xml package from Miscrosoft, which already comes with this "feature".

Before you can start testing the code, you will have to create the file "template.xlsx". 
This file will be the base of the generated xlsx.

Excel has a feature called insert names.

[image:DefiningNamesExcel2010.png]

You will have to create the two names, which will be used in the example code: "header" and "sample1".

The below image shows just the creation of the "sample1", you will have to do it the same for the "header".
[image:DefiningNamesCExcel2010.png]

Notice that both "names" are on "Sheet3", sheet which we will be deleted progamatically at the end, since it contains just our "templates" and we don't want to appear in the generated file.
Also notice that "sample1" contains cells, which start with the sign "<" and end with the sing ">". They are this way specially marked and we can replace their content, with values from our code.
It is maybe good to note also that the "names", can be anywhere, at any cell position. 

One last detail that you need to do before you start: format the cell with the value "<value>" as date, otherwise it will show as a number.

[image:Formatting.png]

Also make sure that you save your workbook as "xlsx".
If you are using Office 2003 or 2000, you can still save it as an "xlsx", if you install the following compatibility pack provided by Microsoft:
"Microsoft Office Compatibility Pack for Word, Excel, and PowerPoint File Formats"
[url:http://www.microsoft.com/downloads/en/details.aspx?familyid=941b3470-3ae9-4aee-8f43-c6bb74cd1466&displaylang=en]
 
Now we can proceed with the code.

The below code will create a new xlsx file called "generated.xlsx" based on the xlsx file that we just created. 
More exactly it will copy in "Sheet1" at position "C3" the named range "header", then it will insert just below it the values returned by the method "getSample" using as template the other named range "sample1".

As you can see in the code, you can customize, which parameter marked in the template between "<" and ">" signs will be replaced by what value. 
More exactly the parameter "name" will be replaced by the first item, the "value" with the second one, the "comment" with the third one, from the yield list returned by the method "getSample".  

{code:c#}


        private static String GENERATED_FILE_NAME = @"c:\work\generated.xlsx";
        private static String TEMPLATE_FILE_NAME = @"c:\work\template.xlsx";

        private void button1_Click(object sender, EventArgs e)
        {
            createGeneratedFile();
            openGeneratedFile();
        }

        private IEnumerable<List<object>> getSample()
        {
            var random = new Random();
            
            for (int loop = 0; loop < 3000; loop++)
            {
                yield return new List<object> {"test", DateTime.Now.AddDays(random.NextDouble()*100 - 50), loop};
            }
            
        }

        private void createGeneratedFile()
        {
            using (ExcelHelper helper = new ExcelHelper(TEMPLATE_FILE_NAME, GENERATED_FILE_NAME))
            {
                helper.Direction = ExcelHelper.DirectionType.TOP_TO_DOWN;

                helper.CurrentSheetName = "Sheet1";

                helper.CurrentPosition = new CellRef("C3");

                //the template xlsx should contains the named range "header"; use the command "insert"/"name".
                helper.InsertRange("header");

                //the template xlsx should contains the named range "sample1";
                //inside this range you should have cells with these values:
                //<name> , <value> and <comment>, which will be replaced by the values from the getSample()
                CellRangeTemplate sample1 = helper.CreateCellRangeTemplate("sample1", new List<string> {"name", "value", "comment"}); 
                
                helper.InsertRange(sample1, getSample());
                
                //you could use here other named ranges to insert new cells and call InsertRange as many times you want, 
                //it will be copied one after another;
                //even you can change direction or the current cell/sheet before you insert
                
                //tipically you put all your "template ranges" (the names) on the same sheet and then you just delete it
                helper.DeleteSheet("Sheet3");
            }        
        }

        private void openGeneratedFile()
        {
            System.Diagnostics.Process.Start(GENERATED_FILE_NAME);
        }

{code:c#}

For convenience I've uploaded also the excel file: [file:template.xlsx] and a solution with the sample code: [file:WindowsFormsApplication2.zip].
The template.xlsx I've created (by the way) using google spreadsheets and downloaded as xlsx. :)

Note also that  the method called InsertSheet in class ExcelHelper allows you to clone a "template sheet" and insert it right after the current-sheet (settable by CurrentSheetName); in case you set the CurrentSheetName to null, then it will be inserted in front. After inserting the new sheet, this will become the current sheet. Sample code will follow.
