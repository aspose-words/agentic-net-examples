using System;
using System.IO;
using System.Runtime.InteropServices;

public class Program
{
    public static void Main()
    {
        // Create a temporary text file that will be embedded as an OLE object
        string tempFolder = Path.Combine(Path.GetTempPath(), "OleDemo");
        Directory.CreateDirectory(tempFolder);
        string sourceFilePath = Path.Combine(tempFolder, "SampleDocument.txt");
        File.WriteAllText(sourceFilePath, "This is the original content of the OLE object.");

        // Prepare paths for the Word document
        string wordFilePath = Path.Combine(tempFolder, "OleDemoDocument.docx");

        // Start Word via COM automation
        Type wordType = Type.GetTypeFromProgID("Word.Application");
        if (wordType == null)
        {
            Console.WriteLine("Microsoft Word is not installed on this machine.");
            return;
        }

        dynamic wordApp = null;
        dynamic document = null;
        try
        {
            wordApp = Activator.CreateInstance(wordType);
            wordApp.Visible = false;

            // Add a new document
            document = wordApp.Documents.Add();

            // Insert a paragraph before the OLE object
            dynamic range = document.Range(0, 0);
            range.Text = "Below is an embedded OLE object preserving its original file name and extension:\n";

            // Insert the OLE object (as a package) and set its display label to the original file name
            dynamic oleObject = document.InlineShapes.AddOLEObject(
                ClassType: "Package",
                FileName: sourceFilePath,
                LinkToFile: false,
                DisplayAsIcon: true,
                IconFileName: "",          // Use default icon
                IconIndex: 0,
                IconLabel: Path.GetFileName(sourceFilePath), // Preserve original name
                Range: range
            );

            // Save the document
            document.SaveAs2(wordFilePath);
        }
        finally
        {
            // Clean up COM objects
            if (document != null)
            {
                document.Close(false);
                Marshal.ReleaseComObject(document);
            }
            if (wordApp != null)
            {
                wordApp.Quit();
                Marshal.ReleaseComObject(wordApp);
            }
        }

        // Output the location of the generated files
        Console.WriteLine($"Embedded OLE object created from: {sourceFilePath}");
        Console.WriteLine($"Word document with OLE object saved to: {wordFilePath}");

        // Optional cleanup (comment out if you want to inspect the files)
        //Directory.Delete(tempFolder, true);
    }
}
