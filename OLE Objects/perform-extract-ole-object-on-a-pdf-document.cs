using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractOleFromPdf
{
    static void Main()
    {
        // Path to the source PDF document.
        string pdfPath = @"C:\Docs\source.pdf";

        // Directory where extracted OLE objects will be saved.
        string outputDir = @"C:\Docs\ExtractedOle";
        Directory.CreateDirectory(outputDir);

        // Load the PDF document. No special load options are required for OLE extraction.
        Document doc = new Document(pdfPath);

        // Iterate through all shapes in the document.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Only shapes that contain an OLE object have a non‑null OleFormat.
            OleFormat ole = shape.OleFormat;
            if (ole == null)
                continue;

            // Determine a suitable file name using the suggested extension.
            string extension = ole.SuggestedExtension ?? ".bin";
            string fileName = Path.Combine(outputDir,
                $"OleObject_{shape.GetHashCode()}{extension}");

            // Save the embedded OLE object to the file system.
            // This uses the OleFormat.Save(string) method as required.
            ole.Save(fileName);
        }

        Console.WriteLine("OLE extraction completed.");
    }
}
