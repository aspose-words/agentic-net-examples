using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractOleFromPdf
{
    static void Main()
    {
        // Path to the source PDF document.
        string pdfPath = @"C:\Input\sample.pdf";

        // Load the PDF document. Aspose.Words can import PDF files.
        Document doc = new Document(pdfPath);

        // Directory where extracted OLE objects will be saved.
        string outputDir = @"C:\Output\OleObjects\";
        Directory.CreateDirectory(outputDir);

        // Iterate through all Shape nodes in the document.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Only shapes that contain an OLE object have a non‑null OleFormat.
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue;

            // Determine a file name for the extracted object.
            // Use the suggested extension if available; otherwise default to .bin.
            string extension = string.IsNullOrEmpty(oleFormat.SuggestedExtension)
                ? ".bin"
                : oleFormat.SuggestedExtension;

            // Create a unique file name based on the shape index.
            string fileName = $"OleObject_{shape.GetHashCode()}{extension}";
            string outputPath = Path.Combine(outputDir, fileName);

            // Save the embedded OLE object to the file system.
            // This uses the OleFormat.Save(string) method as required.
            oleFormat.Save(outputPath);

            Console.WriteLine($"Extracted OLE object to: {outputPath}");
        }
    }
}
