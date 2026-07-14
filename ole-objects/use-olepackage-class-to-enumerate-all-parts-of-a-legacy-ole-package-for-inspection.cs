using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare a simple byte array that represents a ZIP file header.
        // In a real scenario this would be the content of the file you want to embed.
        byte[] packageBytes = new byte[] { 0x50, 0x4B, 0x03, 0x04, 0x00, 0x00, 0x00, 0x00 };

        // Insert the OLE package into the document.
        using (MemoryStream packageStream = new MemoryStream(packageBytes))
        {
            // "Package" is the ProgID for a generic OLE package.
            Shape oleShape = builder.InsertOleObject(packageStream, "Package", true, null);

            // Access the OlePackage object to set its properties.
            OlePackage olePackage = oleShape.OleFormat.OlePackage;
            if (olePackage != null)
            {
                olePackage.FileName = "sample.zip";
                olePackage.DisplayName = "Sample Package";
            }
        }

        // Save the document to a temporary file (required for loading later).
        string docPath = Path.Combine(Path.GetTempPath(), "OlePackageDemo.docx");
        doc.Save(docPath);

        // Load the document back (demonstrates loading with default options).
        Document loadedDoc = new Document(docPath);

        // Enumerate all shapes in the document and inspect OLE packages.
        foreach (Shape shape in loadedDoc.GetChildNodes(NodeType.Shape, true))
        {
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue;

            OlePackage package = oleFormat.OlePackage;
            if (package != null)
            {
                Console.WriteLine("Found OLE Package:");
                Console.WriteLine($"  FileName   : {package.FileName}");
                Console.WriteLine($"  DisplayName: {package.DisplayName}");
            }
        }

        // Clean up the temporary file.
        if (File.Exists(docPath))
            File.Delete(docPath);
    }
}
