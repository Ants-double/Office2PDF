using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using PowerPoint= Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Excel;
using System.Runtime.ExceptionServices;
using System.Security;

namespace Office2PDF
{
    internal class OperationOffice
    {

        public static void word2pdf(string sourcePath, string targetPath)
        {
            // Console.WriteLine("hello..");

            Microsoft.Office.Interop.Word.Application myWordApp;
            Microsoft.Office.Interop.Word.Document myWordDoc;
            myWordApp = new Microsoft.Office.Interop.Word.Application();
            // object filepath =  fileString;
            object filepath = sourcePath;
            object oMissing = System.Reflection.Missing.Value;
            myWordDoc = myWordApp.Documents.Open(ref filepath, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            Microsoft.Office.Interop.Word.WdExportFormat paramExportFormat = Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF;
            bool paramOpenAfterExport = false;
            Microsoft.Office.Interop.Word.WdExportOptimizeFor paramExportOptimizeFor =
                    // Word.WdExportOptimizeFor.wdExportOptimizeForPrint;
                    Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen;
            Microsoft.Office.Interop.Word.WdExportRange paramExportRange = Microsoft.Office.Interop.Word.WdExportRange.wdExportAllDocument;
            int paramStartPage = 0;
            int paramEndPage = 0;
            Microsoft.Office.Interop.Word.WdExportItem paramExportItem = Microsoft.Office.Interop.Word.WdExportItem.wdExportDocumentContent;
            bool paramIncludeDocProps = true;
            bool paramKeepIRM = true;
            Microsoft.Office.Interop.Word.WdExportCreateBookmarks paramCreateBookmarks =
                    Microsoft.Office.Interop.Word.WdExportCreateBookmarks.wdExportCreateWordBookmarks;
            bool paramDocStructureTags = true;
            bool paramBitmapMissingFonts = true;
            bool paramUseISO19005_1 = false;
            string paramExportFilePath = targetPath;
            myWordDoc.ExportAsFixedFormat(paramExportFilePath,
                    paramExportFormat, paramOpenAfterExport,
                    paramExportOptimizeFor, paramExportRange, paramStartPage,
                    paramEndPage, paramExportItem, paramIncludeDocProps,
                    paramKeepIRM, paramCreateBookmarks, paramDocStructureTags,
                    paramBitmapMissingFonts, paramUseISO19005_1,
                    ref oMissing);
            myWordDoc.Close(ref oMissing, ref oMissing, ref oMissing);
            myWordDoc = null;
            myWordApp.Quit(ref oMissing, ref oMissing, ref oMissing);
            myWordDoc = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        public static void ppt2pdf(string sourcePath, string targetPath)
        {
            bool result = false;
            object missing = Type.Missing;
            Microsoft.Office.Interop.PowerPoint.Application application = null;
            PowerPoint.Presentation persentation = null;
            try
            {
                application = new PowerPoint.Application();
                persentation = application.Presentations.Open(sourcePath, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
                persentation.SaveAs(targetPath, PowerPoint.PpSaveAsFileType.ppSaveAsPDF, Microsoft.Office.Core.MsoTriState.msoTrue);
                result = true;
            }
            catch
            {
                result = false;
            }
            finally
            {
                if (persentation != null)
                {
                    persentation.Close();
                    persentation = null;
                }
                if (application != null)
                {
                    application.Quit();
                    application = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            // return result;
        }


        [HandleProcessCorruptedStateExceptions]
        [SecurityCritical]
        public static bool ExportWorkbookToPdf(string workbookPath, string outputPath)
        {
            // If either required string is null or empty, stop and bail out
            if (string.IsNullOrEmpty(workbookPath) || string.IsNullOrEmpty(outputPath))
            {
                return false;
            }

            // Create COM Objects
            Application excelApplication;
            Workbook excelWorkbook;

            // Create new instance of Excel
            excelApplication = new Application();

            // Make the process invisible to the user
            excelApplication.ScreenUpdating = false;

            // Make the process silent
            excelApplication.DisplayAlerts = false;

            // Open the workbook that you wish to export to PDF
            excelWorkbook = excelApplication.Workbooks.Open(workbookPath);

            // If the workbook failed to open, stop, clean up, and bail out
            if (excelWorkbook == null)
            {
                excelApplication.Quit();

                excelApplication = null;
                excelWorkbook = null;

                return false;
            }

            var exportSuccessful = true;
            try
            {
                //excelApplication.PrintCommunication = false;
                foreach (Worksheet sheet in excelWorkbook.Worksheets)
                {
                    PageSetup setup = sheet.PageSetup;
                    setup.Zoom = false;
                    setup.FitToPagesWide = 1;
                    setup.FitToPagesTall = false;
                }
                //excelApplication.PrintCommunication = true;
                // Call Excel's native export function (valid in Office 2007 and Office 2010, AFAIK)
                excelWorkbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, outputPath);
            }
            catch (System.Exception ex)
            {
                // Mark the export as failed for the return value...
                exportSuccessful = false;

                // Do something with any exceptions here, if you wish...
                // MessageBox.Show...        
            }
            finally
            {
                // Close the workbook, quit the Excel, and clean up regardless of the results...
                excelWorkbook.Close();
                excelApplication.Quit();

                excelApplication = null;
                excelWorkbook = null;
            }

            // You can use the following method to automatically open the PDF after export if you wish
            // Make sure that the file actually exists first...
            // if (System.IO.File.Exists(outputPath))
            // {
            //     System.Diagnostics.Process.Start(outputPath);
            // }

            return exportSuccessful;
        }


    }


}
