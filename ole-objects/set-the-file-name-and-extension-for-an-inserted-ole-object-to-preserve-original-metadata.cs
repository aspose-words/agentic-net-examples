using System;
using System.IO;
using System.Runtime.InteropServices;

public class OleObjectExample
{
    public static void Main()
    {
        // Create a temporary text file to embed as OLE object
        string tempDir = Path.Combine(Path.GetTempPath(), "OleObjectExample");
        Directory.CreateDirectory(tempDir);
        string sourceFilePath = Path.Combine(tempDir, "SampleDocument.txt");
        File.WriteAllText(sourceFilePath, "This is the original content of the embedded file.");

        // Path for the resulting Word document
        string docPath = Path.Combine(tempDir, "EmbeddedOleObject.docx");

        dynamic wordApp = null;
        dynamic doc = null;

        try
        {
            // Start Word application via COM late binding (no compile‑time reference needed)
            Type wordType = Type.GetTypeFromProgID("Word.Application");
            if (wordType == null)
            {
                Console.WriteLine("Microsoft Word is not installed on this machine. The example cannot run.");
                return;
            }

            wordApp = Activator.CreateInstance(wordType);
            wordApp.Visible = false;

            // Add a new document
            doc = wordApp.Documents.Add();

            // Insert the OLE object (the text file) as an icon with original file name as label
            // Use Type.Missing for optional parameters we don't set
            doc.InlineShapes.AddOLEObject(
                ClassType: Type.Missing,
                FileName: sourceFilePath,
                LinkToFile: false,
                DisplayAsIcon: true,
                IconFileName: Type.Missing,
                IconIndex: 0,
                IconLabel: Path.GetFileName(sourceFilePath) // Preserve original file name and extension
            );

            // Optional: Add some surrounding text
            dynamic range = doc.Content;
            range.InsertBefore("Below is an embedded OLE object preserving its original metadata:\n\n");
            // wdCollapseEnd = 0
            range.Collapse(0);
            range.InsertAfter("\nEnd of document.");

            // Save the document
            doc.SaveAs2(docPath);
        }
        finally
        {
            // Clean up COM objects
            if (doc != null)
            {
                // wdDoNotSaveChanges = 0
                doc.Close(0);
                Marshal.FinalReleaseComObject(doc);
            }

            if (wordApp != null)
            {
                wordApp.Quit();
                Marshal.FinalReleaseComObject(wordApp);
            }

            // Delete the temporary source file (the OLE object file)
            if (File.Exists(sourceFilePath))
            {
                File.Delete(sourceFilePath);
            }
        }

        // Output the location of the generated document (non‑interactive)
        Console.WriteLine($"Word document with embedded OLE object created at: {docPath}");
    }
}
