using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a source document with sample paragraphs.
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("First paragraph.");
        srcBuilder.Writeln("Second paragraph.");
        srcBuilder.Writeln("Third paragraph.");
        sourceDoc.Save("source.docx");

        // Load the source document.
        Document loadedSource = new Document("source.docx");

        // Extract all paragraph nodes from the source document.
        NodeCollection sourceParagraphs = loadedSource.GetChildNodes(NodeType.Paragraph, true);
        if (sourceParagraphs.Count == 0)
            throw new InvalidOperationException("No paragraphs were found in the source document.");

        // Create a new empty destination document and ensure it has the minimal structure.
        Document destDoc = new Document();
        destDoc.RemoveAllChildren();
        destDoc.EnsureMinimum(); // Adds a Section, Body, and an empty Paragraph.

        // Get the body of the destination document.
        Body destBody = destDoc.FirstSection.Body;

        // Prepare an importer to copy nodes while preserving source formatting.
        NodeImporter importer = new NodeImporter(loadedSource, destDoc, ImportFormatMode.KeepSourceFormatting);

        // Prepend the extracted paragraphs to the beginning of the destination body.
        // Iterate in reverse order because PrependChild inserts before existing content.
        for (int i = sourceParagraphs.Count - 1; i >= 0; i--)
        {
            Node importedNode = importer.ImportNode(sourceParagraphs[i], true);
            destBody.PrependChild(importedNode);
        }

        // Save the resulting document.
        destDoc.Save("result.docx");

        // Verify that the output file was created.
        if (!File.Exists("result.docx"))
            throw new InvalidOperationException("The result document was not created.");

        Console.WriteLine("Document created successfully.");
    }
}
