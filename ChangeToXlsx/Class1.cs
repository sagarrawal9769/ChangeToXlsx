using System;
using Microsoft.Office.Interop.Excel;
using System.Activities;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace ConvertToXlsx
{
    public class ConvertXlsToXlsx : CodeActivity
    {

        // Initializing In Or Out Arguments
        [RequiredArgument]
        [Category("Input")]
        [DisplayName("Input File Path")]
        [Description("Provide input file path")]
        public InArgument<string> InputFilePath { get; set; }

        [RequiredArgument]
        [Category("Input")]
        [DisplayName("Output Directory Path")]
        [Description("Provide Output Directory Path")]
        public InArgument<string> OutputFilePath { get; set; }

        [RequiredArgument]
        [Category("Input")]
        [DisplayName("Output File Name")]
        [Description("Provide Output File Name (without extension)")]
        public InArgument<string> OutputFileName { get; set; }

        protected override void Execute(CodeActivityContext context)
        {

            // Getting Data From Input 
            string inputFile = InputFilePath.Get(context);
            string outputFile = OutputFilePath.Get(context);
            string outputFileName = OutputFileName.Get(context);

            // Making sure to close any previously opened Instance
            Application excelApp = null;
            Workbook workbook = null;

            try
            {
                //Creating a new instance of Excel Application
                excelApp = new Application();
                //Creating a new instance of the Excel Workbook
                workbook = excelApp.Workbooks.Open(inputFile, ReadOnly: true);

                //Check if the file format is .xslx or not 
                if (workbook.FileFormat != XlFileFormat.xlOpenXMLWorkbook)
                {
                    //if not convert them to .xslx
                    workbook.SaveAs(outputFile + outputFileName, XlFileFormat.xlOpenXMLWorkbook);
                }
                else
                {
                    //if yes then also convert them to .xslx for removing any chances of conversion errors in future
                    workbook.SaveAs(outputFile + outputFileName, XlFileFormat.xlOpenXMLWorkbook);
                }
                //closing the workbook while making sure not to make any changes to orignal file 
                workbook.Close(false);
                //closing the connection of the workbook with the excel application
                excelApp.Quit();
            }
            catch (Exception ex)
            {
                throw new Exception($"Error converting {inputFile} to {outputFile}: {ex.Message}");
            }
            finally
            {
                // stoping/closing the instance of workbook
                if (workbook != null)
                {
                    Marshal.FinalReleaseComObject(workbook);
                    workbook = null;
                }
                // stoping/closing the instance of excell application
                if (excelApp != null)
                {
                    Marshal.FinalReleaseComObject(excelApp);
                    excelApp = null;
                }

                // Realisng any memory that is no longer in use
                GC.WaitForPendingFinalizers();
                GC.Collect();
               
               
            }
        }
    }
}
