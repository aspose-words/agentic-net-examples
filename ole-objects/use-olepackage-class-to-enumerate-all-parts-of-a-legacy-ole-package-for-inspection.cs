using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Path for the temporary document.
        string docPath = Path.Combine(Directory.GetCurrentDirectory(), "OlePackageExample.docx");

        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare some sample data to embed as an OLE package.
        byte[] sampleData = System.Text.Encoding.UTF8.GetBytes("This is the content of the embedded file.");

        // Insert the OLE package into the document.
        using (MemoryStream dataStream = new MemoryStream(sampleData))
        {
            // Insert as an OLE object with progId "Package".
            Shape oleShape = builder.InsertOleObject(dataStream, "Package", true, null);

            // Access the OlePackage and set its properties.
            OlePackage olePackage = oleShape.OleFormat.OlePackage;
            olePackage.FileName = "Sample.txt";
            olePackage.DisplayName = "Sample Display.txt";
        }

        // Save the document.
        doc.Save(docPath);

        // Load the document back.
        Document loadedDoc = new Document(docPath);

        // Enumerate all shapes and inspect OLE packages.
        foreach (Shape shape in loadedDoc.GetChildNodes(NodeType.Shape, true))
        {
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat != null && oleFormat.OlePackage != null)
            {
                OlePackage pkg = oleFormat.OlePackage;
                Console.WriteLine("Found OLE Package:");
                Console.WriteLine($"  FileName   : {pkg.FileName}");
                Console.WriteLine($"  DisplayName: {pkg.DisplayName}");
            }
        }
    }
}
