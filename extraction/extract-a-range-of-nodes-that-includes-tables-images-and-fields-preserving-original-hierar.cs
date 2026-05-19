using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;

public class ExtractRangeExample
{
    public static void Main()
    {
        // Create a sample source document containing a field, a table with an image, and surrounding paragraphs.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Intro paragraph.
        builder.Writeln("Intro paragraph before the extracted range.");

        // Paragraph that contains a MERGEFIELD.
        builder.InsertField(" MERGEFIELD SampleField ");

        // Table with one cell that holds an image.
        builder.StartTable();
        builder.InsertCell();

        // Insert a tiny PNG image from a Base64 string.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=");
        using (MemoryStream imgStream = new MemoryStream(pngBytes))
        {
            builder.InsertImage(imgStream);
        }

        builder.EndRow();
        builder.EndTable();

        // Closing paragraph.
        builder.Writeln("Closing paragraph after the extracted range.");

        // Save the source document (optional, just for verification).
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // Load the document back (simulating a real‑world scenario).
        Document loadedDoc = new Document(sourcePath);

        // Identify the start node: the paragraph that contains the first field.
        Field firstField = loadedDoc.Range.Fields[0];
        Paragraph startParagraph = (Paragraph)firstField.Start.GetAncestor(NodeType.Paragraph);
        if (startParagraph == null)
            throw new InvalidOperationException("Start paragraph not found.");

        // Identify the end node: the first table in the document.
        Table endTable = loadedDoc.GetChildNodes(NodeType.Table, true)[0] as Table;
        if (endTable == null)
            throw new InvalidOperationException("Table not found.");

        // Prepare the destination document.
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren(); // Clear the default empty section/paragraph.
        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);
        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // Use NodeImporter to import nodes while preserving formatting.
        NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);

        // Import the start paragraph (contains the field) and append it.
        Node importedStart = importer.ImportNode(startParagraph, true);
        resultBody.AppendChild(importedStart);

        // Import the table (contains the image) and append it.
        Node importedTable = importer.ImportNode(endTable, true);
        resultBody.AppendChild(importedTable);

        // Save the extracted range to a new document.
        const string resultPath = "extracted_range.docx";
        resultDoc.Save(resultPath);

        // Verify that the output file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The extracted document was not created.");

        // Confirmation message.
        Console.WriteLine("Extraction completed successfully. Output file: " + resultPath);
    }
}
