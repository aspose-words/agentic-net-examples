using System;
using System.IO;
using System.Runtime.InteropServices;

public class Program
{
    public static void Main()
    {
        // Create a temporary text file to embed.
        string tempDir = Path.GetTempPath();
        string sampleFile = Path.Combine(tempDir, "Sample.txt");
        File.WriteAllText(sampleFile, "This is sample content for OLE object.");

        // Start Word via COM.
        Type wordType = Type.GetTypeFromProgID("Word.Application");
        if (wordType == null)
        {
            Console.WriteLine("Microsoft Word is not installed.");
            return;
        }

        dynamic wordApp = Activator.CreateInstance(wordType);
        try
        {
            wordApp.Visible = false;
            dynamic documents = wordApp.Documents;
            dynamic doc = documents.Add();

            // Insert the OLE object as an icon using the default system icon.
            dynamic range = doc.Range(0, 0);
            range.InlineShapes.AddOLEObject(
                null,                 // ClassType
                sampleFile,           // FileName
                false,                // LinkToFile
                true,                 // DisplayAsIcon
                Type.Missing,         // IconFileName (default)
                Type.Missing,         // IconIndex (default)
                "Sample Text File",   // IconLabel
                Type.Missing          // Range (not used here)
            );

            // Save the document.
            string docPath = Path.Combine(tempDir, "OleIconDemo.docx");
            doc.SaveAs2(docPath);
            doc.Close();

            Console.WriteLine($"Document saved to: {docPath}");
        }
        finally
        {
            // Quit Word and release COM object.
            wordApp.Quit();
            Marshal.ReleaseComObject(wordApp);
        }
    }
}
