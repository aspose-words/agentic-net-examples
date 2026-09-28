using System;
using System.IO;
using System.Runtime.InteropServices;

public class Program
{
    public static void Main()
    {
        // Folder containing Word documents
        string folderPath = Path.Combine(Directory.GetCurrentDirectory(), "WordDocs");
        // Path to the Excel file to embed
        string excelPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.xlsx");

        // Ensure the folder exists
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Ensure the Excel file exists
        if (!File.Exists(excelPath))
        {
            Console.WriteLine($"Excel file not found: {excelPath}");
            return;
        }

        // Create Word Application via late binding (no compile‑time reference to Interop)
        Type wordAppType = Type.GetTypeFromProgID("Word.Application");
        if (wordAppType == null)
        {
            Console.WriteLine("Microsoft Word is not installed on this machine.");
            return;
        }

        dynamic wordApp = null;
        try
        {
            wordApp = Activator.CreateInstance(wordAppType);
            wordApp.Visible = false;

            foreach (string docPath in Directory.GetFiles(folderPath, "*.docx"))
            {
                dynamic doc = null;
                try
                {
                    // Open the document (ReadOnly = false, Visible = false)
                    object readOnly = false;
                    object isVisible = false;
                    object missing = Type.Missing;

                    doc = wordApp.Documents.Open(
                        docPath,
                        ref missing,          // ConfirmConversions
                        ref readOnly,         // ReadOnly
                        ref missing,          // AddToRecentFiles
                        ref missing,          // PasswordDocument
                        ref missing,          // PasswordTemplate
                        ref missing,          // Revert
                        ref missing,          // WritePasswordDocument
                        ref missing,          // WritePasswordTemplate
                        ref missing,          // Format
                        ref missing,          // Encoding
                        ref isVisible,        // Visible
                        ref missing,          // OpenAndRepair
                        ref missing,          // DocumentDirection
                        ref missing,          // NoEncodingDialog
                        ref missing);         // XMLTransform

                    // Get a range at the end of the document
                    dynamic range = doc.Content;
                    // wdCollapseEnd = 0
                    range.Collapse(0);

                    // Prepare parameters for AddOLEObject
                    object classType = "Excel.Sheet";
                    object fileName = excelPath;
                    object linkToFile = false;
                    object displayAsIcon = false;
                    object iconFileName = missing;
                    object iconIndex = missing;
                    object iconLabel = missing;
                    object oleRange = range;

                    // Insert the Excel OLE object
                    dynamic ole = doc.InlineShapes.AddOLEObject(
                        classType,
                        fileName,
                        linkToFile,
                        displayAsIcon,
                        iconFileName,
                        iconIndex,
                        iconLabel,
                        oleRange);

                    // Save and close the document
                    doc.Save();
                }
                finally
                {
                    if (doc != null)
                    {
                        doc.Close();
                        Marshal.FinalReleaseComObject(doc);
                    }
                }
            }
        }
        finally
        {
            if (wordApp != null)
            {
                wordApp.Quit();
                Marshal.FinalReleaseComObject(wordApp);
            }
        }

        Console.WriteLine("Processing completed.");
    }
}
