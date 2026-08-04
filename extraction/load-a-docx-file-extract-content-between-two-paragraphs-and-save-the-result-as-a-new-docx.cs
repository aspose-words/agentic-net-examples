using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample source document with four paragraphs.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");
        builder.Writeln("Paragraph 3");
        builder.Writeln("Paragraph 4");
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // Load the document from the file system.
        Document loadedDoc = new Document(sourcePath);

        // Identify the start and end paragraphs (inclusive extraction).
        Paragraph startParagraph = loadedDoc.FirstSection.Body.Paragraphs[1];
        Paragraph endParagraph = loadedDoc.FirstSection.Body.Paragraphs[2];

        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Required paragraphs were not found.");

        // Prepare the destination document.
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren(); // Clear the default nodes.

        // Create a new section and body for the result document.
        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);
        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // Import the selected paragraphs into the result document.
        NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedStart = importer.ImportNode(startParagraph, true);
        Node importedEnd = importer.ImportNode(endParagraph, true);

        resultBody.AppendChild(importedStart);
        resultBody.AppendChild(importedEnd);

        // Save the extracted content as a new DOCX file.
        const string resultPath = "extracted.docx";
        resultDoc.Save(resultPath);

        // Verify that the output file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The extracted document was not created.");
    }
}
