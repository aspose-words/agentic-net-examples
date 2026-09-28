using System;
using System.IO;
using System.Runtime.InteropServices;

public class OleObjectExample
{
    public static void Main()
    {
        // Create temporary text file to embed as OLE object
        string tempTextPath = Path.Combine(Path.GetTempPath(), "OleSample.txt");
        File.WriteAllText(tempTextPath, "This is sample text for an OLE object.");

        // Path for the generated Word document
        string docPath = Path.Combine(Path.GetTempPath(), "OleExample.docx");

        dynamic wordApp = null;
        dynamic doc = null;
        dynamic inlineShape = null;

        try
        {
            // Start Word application (invisible)
            Type wordType = Type.GetTypeFromProgID("Word.Application");
            wordApp = Activator.CreateInstance(wordType);
            wordApp.Visible = false;

            // Add a new document
            doc = wordApp.Documents.Add();

            // Insert the OLE object (the text file) into the document
            // Parameters: ClassType, FileName, LinkToFile, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Range
            inlineShape = doc.InlineShapes.AddOLEObject(
                "Package",                 // ClassType
                tempTextPath,              // FileName
                false,                     // LinkToFile
                false,                     // DisplayAsIcon
                Type.Missing,              // IconFileName
                Type.Missing,              // IconIndex
                Type.Missing,              // IconLabel
                Type.Missing               // Range
            );

            // Retrieve original dimensions (points)
            double originalWidth = inlineShape.Width;
            double originalHeight = inlineShape.Height;
            Console.WriteLine($"Original Width: {originalWidth} pt, Height: {originalHeight} pt");

            // Adjust size (double the dimensions)
            inlineShape.Width = originalWidth * 2;
            inlineShape.Height = originalHeight * 2;
            Console.WriteLine($"Adjusted Width: {inlineShape.Width} pt, Height: {inlineShape.Height} pt");

            // Save the document
            doc.SaveAs2(docPath);
        }
        finally
        {
            // Clean up COM objects
            if (inlineShape != null) Marshal.ReleaseComObject(inlineShape);
            if (doc != null) Marshal.ReleaseComObject(doc);
            if (wordApp != null)
            {
                wordApp.Quit();
                Marshal.ReleaseComObject(wordApp);
            }

            // Delete temporary files
            if (File.Exists(tempTextPath))
                File.Delete(tempTextPath);
        }

        // Optionally, delete the generated document (uncomment if desired)
        // if (File.Exists(docPath))
        //     File.Delete(docPath);
    }
}
