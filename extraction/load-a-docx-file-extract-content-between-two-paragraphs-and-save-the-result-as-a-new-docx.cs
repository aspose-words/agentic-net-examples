using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample source document with four paragraphs and save it.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");
        builder.Writeln("Paragraph 3");
        builder.Writeln("Paragraph 4");

        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the source document.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // -----------------------------------------------------------------
        // 3. Locate the start and end paragraphs (inclusive boundaries).
        // -----------------------------------------------------------------
        Paragraph startParagraph = loadedDoc.FirstSection.Body.Paragraphs[1]; // "Paragraph 2"
        Paragraph endParagraph   = loadedDoc.FirstSection.Body.Paragraphs[2]; // "Paragraph 3"

        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Boundary paragraphs not found.");

        // -----------------------------------------------------------------
        // 4. Prepare a new empty document that will hold the extracted range.
        // -----------------------------------------------------------------
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren();                     // clear the default section/paragraph
        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);
        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // -----------------------------------------------------------------
        // 5. Import the selected paragraphs into the result document.
        //    Nodes must be imported before they can be inserted into another document.
        // -----------------------------------------------------------------
        NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);

        Node importedStart = importer.ImportNode(startParagraph, true);
        Node importedEnd   = importer.ImportNode(endParagraph,   true);

        resultBody.AppendChild((Paragraph)importedStart);
        resultBody.AppendChild((Paragraph)importedEnd);

        // -----------------------------------------------------------------
        // 6. Save the extracted content as a new DOCX file.
        // -----------------------------------------------------------------
        const string resultPath = "extracted.docx";
        resultDoc.Save(resultPath);

        // -----------------------------------------------------------------
        // 7. Verify that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The extracted document was not created.");

        // Optional: display a confirmation message.
        Console.WriteLine($"Extraction completed successfully. Output saved to '{resultPath}'.");
    }
}
