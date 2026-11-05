using System;
using System.IO;
using System.Runtime.InteropServices;

public class Program
{
    public static void Main()
    {
        string baseDir = Directory.GetCurrentDirectory();
        string docsFolder = Path.Combine(baseDir, "Docs");
        Directory.CreateDirectory(docsFolder);

        string excelPath = Path.Combine(baseDir, "Sample.xlsx");
        if (!File.Exists(excelPath))
        {
            // Create a minimal empty Excel file (just an empty zip structure)
            File.WriteAllBytes(excelPath, new byte[] { });
        }

        // Ensure there is at least one Word document to process
        string[] existingDocs = Directory.GetFiles(docsFolder, "*.docx");
        if (existingDocs.Length == 0)
        {
            CreateSampleWordDocument(Path.Combine(docsFolder, "Document1.docx"));
            CreateSampleWordDocument(Path.Combine(docsFolder, "Document2.docx"));
        }

        // Start Word application via COM
        Type wordType = Type.GetTypeFromProgID("Word.Application");
        if (wordType == null)
        {
            Console.WriteLine("Microsoft Word is not installed on this machine.");
            return;
        }

        dynamic wordApp = null;
        try
        {
            wordApp = Activator.CreateInstance(wordType);
            wordApp.Visible = false;

            string[] docFiles = Directory.GetFiles(docsFolder, "*.docx");
            foreach (string docPath in docFiles)
            {
                dynamic document = null;
                try
                {
                    // Open the document (ReadOnly: false, Visible: false)
                    document = wordApp.Documents.Open(docPath, ReadOnly: false, Visible: false);

                    // Get the end of the document content
                    dynamic range = document.Content;
                    // Collapse to the end (wdCollapseEnd = 0)
                    range.Collapse(0);

                    // Insert the Excel OLE object
                    range.InlineShapes.AddOLEObject(
                        Class: "Excel.Sheet",
                        FileName: excelPath,
                        LinkToFile: false,
                        DisplayAsIcon: false);

                    // Save and close the document
                    document.Save();
                    document.Close();
                }
                finally
                {
                    if (document != null && Marshal.IsComObject(document))
                    {
                        Marshal.FinalReleaseComObject(document);
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
    }

    private static void CreateSampleWordDocument(string filePath)
    {
        Type wordType = Type.GetTypeFromProgID("Word.Application");
        if (wordType == null) return;

        dynamic wordApp = null;
        dynamic doc = null;
        try
        {
            wordApp = Activator.CreateInstance(wordType);
            wordApp.Visible = false;
            doc = wordApp.Documents.Add();
            doc.Content.Text = "Sample document created for OLE insertion test.";
            doc.SaveAs2(filePath);
            doc.Close();
        }
        finally
        {
            if (doc != null && Marshal.IsComObject(doc))
                Marshal.FinalReleaseComObject(doc);
            if (wordApp != null)
            {
                wordApp.Quit();
                Marshal.FinalReleaseComObject(wordApp);
            }
        }
    }
}
