using System;
using System.Linq;
using System.Activities;
using System.Activities.Statements;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Words;

namespace ASPOSETEST
{

    class Program
    {
        static Program()
        {
            CosturaUtility.Initialize();
        }
        private  int totalfilecount = 0;
        private string author, lastsavedby;

        private void updateProperties(string dirFile)
        {
            foreach (string fileName in Directory.GetFiles(dirFile))
            {
                if (Path.GetExtension(fileName) == ".xlsx" || Path.GetExtension(fileName) == ".XLSX"
                    || Path.GetExtension(fileName) == ".XLS"
                    || Path.GetExtension(fileName) == ".xls"
                    || Path.GetExtension(fileName) == ".xlsm"
                    || Path.GetExtension(fileName) == ".XLSM")
                {
                    try
                    {
                        Workbook workbook = new Workbook(fileName);
                        //Get built-in properties of the excel file in a collection object
                        Aspose.Cells.Properties.DocumentPropertyCollection builtInProperties = workbook.Worksheets.BuiltInDocumentProperties;

                        //Get individual property
                        Aspose.Cells.Properties.DocumentProperty propertyAuthor = builtInProperties["Author"];
                        Aspose.Cells.Properties.DocumentProperty propertylastSavedByAuthor = builtInProperties["LastSavedBy"];

                        //Change individual property value as required
                        propertyAuthor.Value = author;
                        propertylastSavedByAuthor.Value = lastsavedby;

                        //Save the modified workbook
                        workbook.Save(fileName);
                        totalfilecount++;
                        Console.WriteLine(Path.GetFileName(fileName) + " is changed.\n" + "{ Author: " + author + ", LastSavedBy: " + lastsavedby+" }\n\n");
                    }
                    catch (Exception)
                    {

                        Console.WriteLine(Path.GetFileName(fileName) + " is not a genuine Spreadsheet file / my have been corrupted.\n");
                    }
                }
                else if (Path.GetExtension(fileName) == ".docx" || Path.GetExtension(fileName) == ".DOCX" ||
                    Path.GetExtension(fileName) == ".DOC" || Path.GetExtension(fileName) == ".doc")
                {
                    try
                    {
                        Document doc = new Document(fileName);

                        //Get built-in properties of the excel file in a collection object
                        Aspose.Words.Properties.DocumentPropertyCollection builtInProperties = doc.BuiltInDocumentProperties;

                        //Get individual property
                        Aspose.Words.Properties.DocumentProperty propertyAuthor = builtInProperties["Author"];
                        Aspose.Words.Properties.DocumentProperty propertylastSavedByAuthor = builtInProperties["LastSavedBy"];

                        //Change individual property value as required
                        propertyAuthor.Value = author;
                        propertylastSavedByAuthor.Value = lastsavedby;

                        //Save the modified workbook
                        doc.Save(fileName);
                        totalfilecount++;
                        Console.WriteLine(Path.GetFileName(fileName) + " is changed.\n" + "{ Author: " + author + ", LastSavedBy: " + lastsavedby + " }\n\n");
                    }
                    catch (Exception)
                    {
                        Console.WriteLine(Path.GetFileName(fileName) + " is not a genuine Document file / my have been corrupted.\n");
                    }
                }
            }
        }
        private void REC(string directory)
        {
            if (directory == null) return;
            foreach (string dirFile in Directory.GetDirectories(directory))
            {
                REC(dirFile);
            }
            updateProperties(directory);
            return;

        }
        static void Main(string[] args)
        {
            Activity workflow1 = new Workflow1();
            string path = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
            Program obj = new Program();
            Console.Write("Enter Author: ");
            obj.author = Console.ReadLine();
            Console.Write("Enter LastSavedBy: ");
            obj.lastsavedby = Console.ReadLine();
            obj.REC(path);
            Console.WriteLine("Successfully Completed.\nTotal file changed : " + obj.totalfilecount+"\nPlease press enter to exit . . . . .");
            WorkflowInvoker.Invoke(workflow1);
            Console.ReadLine();
        }
    }
}
