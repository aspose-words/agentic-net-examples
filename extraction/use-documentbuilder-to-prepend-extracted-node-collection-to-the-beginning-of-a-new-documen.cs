using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // ------------------------------------------------------------
        // Create a source document with sample paragraphs.
        // ------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("Paragraph A");
        srcBuilder.Writeln("Paragraph B");
        srcBuilder.Writeln("Paragraph C");
        sourceDoc.Save("source.docx");

        // ------------------------------------------------------------
        // Load the source document (already in memory) and collect its paragraphs.
        // ------------------------------------------------------------
        ParagraphCollection sourceParagraphs = sourceDoc.FirstSection.Body.Paragraphs;
        if (sourceParagraphs.Count == 0)
            throw new InvalidOperationException("No paragraphs found in the source document.");

        // ------------------------------------------------------------
        // Create a new empty document with a clean Section/Body hierarchy.
        // ------------------------------------------------------------
        Document newDoc = new Document();
        newDoc.RemoveAllChildren(); // Remove the default empty section.

        Section section = new Section(newDoc);
        newDoc.AppendChild(section);
        Body body = new Body(newDoc);
        section.AppendChild(body);

        // ------------------------------------------------------------
        // Add an initial paragraph to the new document.
        // ------------------------------------------------------------
        DocumentBuilder newBuilder = new DocumentBuilder(newDoc);
        newBuilder.Writeln("Original content");
        Paragraph originalParagraph = newDoc.FirstSection.Body.LastParagraph;

        // ------------------------------------------------------------
        // Prepend each extracted paragraph before the original paragraph.
        // Use NodeImporter to import nodes from the source document into the destination document.
        // ------------------------------------------------------------
        foreach (Paragraph srcPara in sourceParagraphs)
        {
            // Import the source paragraph into the destination document, preserving formatting.
            NodeImporter importer = new NodeImporter(srcPara.Document, newDoc, ImportFormatMode.KeepSourceFormatting);
            Node importedNode = importer.ImportNode(srcPara, true);

            // The imported node is a Paragraph; insert it before the original paragraph.
            body.InsertBefore(importedNode, originalParagraph);
        }

        // ------------------------------------------------------------
        // Save the resulting document and verify its creation.
        // ------------------------------------------------------------
        newDoc.Save("result.docx");

        if (!File.Exists("result.docx"))
            throw new InvalidOperationException("The result document was not created.");
    }
}
