using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

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

        // Load the source document (demonstrates loading workflow).
        Document loadedSource = new Document("source.docx");

        // Extract all paragraph nodes from the source document.
        NodeCollection sourceParagraphs = loadedSource.FirstSection.Body.GetChildNodes(NodeType.Paragraph, true);
        if (sourceParagraphs.Count == 0)
            throw new InvalidOperationException("No paragraphs were extracted from the source document.");

        // Create a new destination document.
        Document destDoc = new Document();

        // Prepare an importer to copy nodes while preserving source formatting.
        NodeImporter importer = new NodeImporter(loadedSource, destDoc, ImportFormatMode.KeepSourceFormatting);

        // Get the body of the destination document where paragraphs will be prepended.
        Body destBody = destDoc.FirstSection.Body;

        // Prepend the extracted paragraphs to the beginning of the destination document.
        // Iterate in reverse order and prepend each imported paragraph.
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

        Console.WriteLine("Extraction and prepend operation completed successfully.");
    }
}
