using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // ------------------------------------------------------------
        // 1. Create a sample source document with identifiable boundaries.
        // ------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Paragraph before start.");
        // Start boundary.
        builder.Writeln("Start paragraph.");
        // Content to be extracted.
        builder.Writeln("Extracted paragraph 1.");
        builder.Writeln("Extracted paragraph 2.");
        // End boundary.
        builder.Writeln("End paragraph.");
        builder.Writeln("Paragraph after end.");

        // ------------------------------------------------------------
        // 2. Locate the start and end paragraphs (inclusive boundaries).
        // ------------------------------------------------------------
        Paragraph startPara = sourceDoc.FirstSection.Body.Paragraphs[1];
        Paragraph endPara = sourceDoc.FirstSection.Body.Paragraphs[4];

        if (startPara == null || endPara == null)
            throw new InvalidOperationException("Boundary paragraphs not found.");

        // ------------------------------------------------------------
        // 3. Prepare a new empty document that will hold the extracted range.
        // ------------------------------------------------------------
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren(); // Remove the default empty section.

        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);

        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // ------------------------------------------------------------
        // 4. Import the paragraphs between the boundaries (inclusive) into the new document.
        // ------------------------------------------------------------
        // Use a NodeImporter to transfer nodes between documents safely.
        NodeImporter importer = new NodeImporter(sourceDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);

        // Determine the indices of the start and end paragraphs.
        int startIndex = sourceDoc.FirstSection.Body.Paragraphs.IndexOf(startPara);
        int endIndex = sourceDoc.FirstSection.Body.Paragraphs.IndexOf(endPara);

        if (startIndex < 0 || endIndex < 0 || endIndex < startIndex)
            throw new InvalidOperationException("Invalid paragraph range.");

        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph srcParagraph = sourceDoc.FirstSection.Body.Paragraphs[i];
            // Import the paragraph into the destination document.
            Node importedNode = importer.ImportNode(srcParagraph, true);
            resultBody.AppendChild(importedNode);
        }

        // ------------------------------------------------------------
        // 5. Verify that something was extracted.
        // ------------------------------------------------------------
        if (resultBody.Paragraphs.Count == 0)
            throw new InvalidOperationException("No content was extracted.");

        // ------------------------------------------------------------
        // 6. Save the extracted document encrypted with a password.
        // ------------------------------------------------------------
        string outputPath = "extracted-encrypted.docx";
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            Password = "Secret123",
            SaveFormat = SaveFormat.Docx
        };
        resultDoc.Save(outputPath, saveOptions);

        // ------------------------------------------------------------
        // 7. Validate that the encrypted file was created.
        // ------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Encrypted output file was not created.");

        Console.WriteLine("Extraction and encryption completed.");
    }
}
