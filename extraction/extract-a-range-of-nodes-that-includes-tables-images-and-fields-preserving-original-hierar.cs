using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a source document and populate it with a field, a table, an image and some text.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Paragraph containing a MERGEFIELD.
        builder.Writeln("Paragraph before field.");
        builder.InsertField("MERGEFIELD SampleField");
        builder.Writeln(); // End of field paragraph.

        // Insert a simple 2x2 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();
        builder.EndTable();

        // Insert a tiny PNG image (1x1 pixel) from a byte array.
        const string pngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ/2ZL2AAAAAElFTkSuQmCC";
        byte[] imageBytes = Convert.FromBase64String(pngBase64);
        builder.InsertImage(imageBytes);
        builder.Writeln("Paragraph after image.");

        // -----------------------------------------------------------------
        // Extract the range that includes the field, table, and image.
        // For this example we extract the entire body of the first section,
        // which contains all the nodes we just added.
        // -----------------------------------------------------------------
        Body sourceBody = sourceDoc.FirstSection.Body;

        // Create a destination document with an empty structure.
        Document destDoc = new Document();
        destDoc.RemoveAllChildren(); // Remove the default section/paragraph.
        Section destSection = new Section(destDoc);
        destDoc.AppendChild(destSection);
        Body destBody = new Body(destDoc);
        destSection.AppendChild(destBody);

        // Use NodeImporter to copy nodes while preserving formatting and hierarchy.
        NodeImporter importer = new NodeImporter(sourceDoc, destDoc, ImportFormatMode.KeepSourceFormatting);

        foreach (Node srcNode in sourceBody)
        {
            // Import each block-level node (Paragraph, Table, etc.) into the destination.
            Node importedNode = importer.ImportNode(srcNode, true);
            destBody.AppendChild(importedNode);
        }

        // Save the extracted document.
        const string outputPath = "Extracted.docx";
        destDoc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file '{outputPath}'.");

        Console.WriteLine($"Extraction completed successfully. Output saved to '{outputPath}'.");
    }
}
