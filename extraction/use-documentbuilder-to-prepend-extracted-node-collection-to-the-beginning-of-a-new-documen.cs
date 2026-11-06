using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // ------------------------------------------------------------
        // Create a sample source document with several paragraphs.
        // ------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("First paragraph in source.");
        srcBuilder.Writeln("Second paragraph in source.");
        srcBuilder.Writeln("Third paragraph in source.");
        sourceDoc.Save("source.docx");

        // ------------------------------------------------------------
        // Load the source document.
        // ------------------------------------------------------------
        Document loadedSource = new Document("source.docx");

        // ------------------------------------------------------------
        // Extract all paragraphs from the source document.
        // ------------------------------------------------------------
        ParagraphCollection sourceParagraphs = loadedSource.FirstSection.Body.Paragraphs;
        if (sourceParagraphs == null || sourceParagraphs.Count == 0)
            throw new InvalidOperationException("No paragraphs were found in the source document.");

        // ------------------------------------------------------------
        // Create a new destination document with an explicit Section/Body.
        // ------------------------------------------------------------
        Document destDoc = new Document();
        destDoc.RemoveAllChildren();

        Section destSection = new Section(destDoc);
        destDoc.AppendChild(destSection);

        Body destBody = new Body(destDoc);
        destSection.AppendChild(destBody);

        // Add an initial paragraph to demonstrate prepend operation.
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Existing paragraph in destination document.");

        // ------------------------------------------------------------
        // Prepare an importer to move nodes from the source document to the destination.
        // ------------------------------------------------------------
        NodeImporter importer = new NodeImporter(loadedSource, destDoc, ImportFormatMode.KeepSourceFormatting);

        // Prepend the extracted paragraphs to the beginning of the destination body.
        foreach (Paragraph para in sourceParagraphs)
        {
            // Import the paragraph into the destination document.
            Node importedNode = importer.ImportNode(para, true);
            if (importedNode == null)
                continue; // Safety check.

            // Insert before the current first node (or append if the body is empty).
            Node firstNode = destBody.FirstChild;
            if (firstNode != null)
                destBody.InsertBefore(importedNode, firstNode);
            else
                destBody.AppendChild(importedNode);
        }

        // ------------------------------------------------------------
        // Save the resulting document.
        // ------------------------------------------------------------
        destDoc.Save("result.docx");

        // Validate that the output file was created.
        if (!File.Exists("result.docx"))
            throw new InvalidOperationException("The result document was not created.");

        Console.WriteLine("Extraction and prepend operation completed successfully.");
    }
}
