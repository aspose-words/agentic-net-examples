using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a sample source document containing paragraphs, a table, an image, and a field.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Paragraph before the range.
        builder.Writeln("Paragraph before the extraction range.");

        // Start marker paragraph (inclusive).
        builder.Writeln("Start of extraction range.");

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
        builder.EndTable();

        // Insert an inline image inside its own paragraph.
        builder.Writeln(); // Ensure a new paragraph for the image.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=");
        using (MemoryStream imgStream = new MemoryStream(pngBytes))
        {
            builder.InsertImage(imgStream);
        }

        // Insert a field (MERGEFIELD) inside its own paragraph.
        builder.Writeln(); // New paragraph for the field.
        builder.InsertField(" MERGEFIELD SampleField ");

        // End marker paragraph (inclusive).
        builder.Writeln("End of extraction range.");

        // Paragraph after the range.
        builder.Writeln("Paragraph after the extraction range.");

        // Save the source document (optional, for verification).
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // Locate the start and end paragraphs that bound the extraction range.
        Paragraph startParagraph = sourceDoc.FirstSection.Body.Paragraphs[1]; // "Start of extraction range."
        Paragraph endParagraph = sourceDoc.FirstSection.Body.Paragraphs[5];   // "End of extraction range."

        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Failed to locate the start or end paragraph for extraction.");

        // Prepare the destination document.
        Document destDoc = new Document();
        destDoc.RemoveAllChildren(); // Clear the default empty section/paragraph.
        Section destSection = new Section(destDoc);
        destDoc.AppendChild(destSection);
        Body destBody = new Body(destDoc);
        destSection.AppendChild(destBody);

        // Import nodes from the source document between the start and end paragraphs (inclusive).
        NodeImporter importer = new NodeImporter(sourceDoc, destDoc, ImportFormatMode.KeepSourceFormatting);
        Node currentNode = startParagraph;
        while (true)
        {
            // Import the current node (deep clone) and append it to the destination body.
            Node importedNode = importer.ImportNode(currentNode, true);
            destBody.AppendChild(importedNode);

            // Break when the end paragraph has been processed.
            if (currentNode == endParagraph)
                break;

            // Move to the next sibling node in the source document.
            currentNode = currentNode.NextSibling;
            if (currentNode == null)
                throw new InvalidOperationException("Reached the end of the document before finding the end marker.");
        }

        // Save the extracted range to a new document.
        const string resultPath = "extracted_range.docx";
        destDoc.Save(resultPath);

        // Validate that the output file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The extracted document was not created.");

        // Optional: output a simple confirmation (no interactive input required).
        Console.WriteLine("Extraction completed successfully.");
    }
}
