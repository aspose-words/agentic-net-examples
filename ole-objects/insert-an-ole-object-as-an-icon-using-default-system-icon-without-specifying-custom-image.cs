using System;
using System.IO;
using System.Runtime.InteropServices;

public class Program
{
    public static void Main()
    {
        // Create a temporary text file to embed
        string tempFile = Path.Combine(Path.GetTempPath(), "Sample.txt");
        File.WriteAllText(tempFile, "This is sample content for OLE object.");

        // Start Word via COM late binding (no compile‑time reference to Interop assembly)
        Type wordProgId = Type.GetTypeFromProgID("Word.Application");
        if (wordProgId == null)
        {
            Console.WriteLine("Microsoft Word is not installed on this machine.");
            return;
        }

        dynamic wordApp = Activator.CreateInstance(wordProgId);
        wordApp.Visible = false;
        dynamic doc = null;

        try
        {
            // Add a new document
            doc = wordApp.Documents.Add();

            // Insert OLE object as an icon using the default system icon
            object missing = Type.Missing;
            doc.InlineShapes.AddOLEObject(
                ClassType: missing,
                FileName: tempFile,
                LinkToFile: false,
                DisplayAsIcon: true,
                IconFileName: missing,
                IconIndex: missing,
                IconLabel: "Sample Text File",
                Range: missing
            );

            // Save the document to a temporary location
            string docPath = Path.Combine(Path.GetTempPath(), "OleExample.docx");
            doc.SaveAs2(docPath);
            Console.WriteLine($"Document saved to: {docPath}");
        }
        finally
        {
            // Clean up Word objects
            if (doc != null)
            {
                doc.Close(false);
                Marshal.ReleaseComObject(doc);
            }

            if (wordApp != null)
            {
                wordApp.Quit();
                Marshal.ReleaseComObject(wordApp);
            }
        }

        // Delete the temporary text file
        if (File.Exists(tempFile))
        {
            File.Delete(tempFile);
        }
    }
}
