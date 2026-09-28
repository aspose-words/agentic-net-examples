using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample source document with several paragraphs.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Paragraph 1 - before");
        builder.Writeln("Paragraph 2 - start");
        builder.Writeln("Paragraph 3 - middle");
        builder.Writeln("Paragraph 4 - end");
        builder.Writeln("Paragraph 5 - after");
        sourceDoc.Save("source.docx");

        // Step 2: Load the created document.
        Document loadedDoc = new Document("source.docx");

        // Step 3: Identify the start and end paragraphs for extraction.
        Paragraph startPara = loadedDoc.FirstSection.Body.Paragraphs[1]; // "Paragraph 2 - start"
        Paragraph endPara = loadedDoc.FirstSection.Body.Paragraphs[3];   // "Paragraph 4 - end"

        if (startPara == null || endPara == null)
            throw new InvalidOperationException("Boundary paragraphs not found.");

        // Step 4: Determine the indices of the start and end paragraphs.
        int startIndex = loadedDoc.FirstSection.Body.Paragraphs.IndexOf(startPara);
        int endIndex = loadedDoc.FirstSection.Body.Paragraphs.IndexOf(endPara);
        if (startIndex < 0 || endIndex < 0 || startIndex > endIndex)
            throw new InvalidOperationException("Invalid paragraph range.");

        // Step 5: Create a new document to hold the extracted content.
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren(); // Remove default empty section.

        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);

        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // Step 6: Import and append each paragraph within the range (inclusive).
        NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);
        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph para = loadedDoc.FirstSection.Body.Paragraphs[i];
            Node importedNode = importer.ImportNode(para, true);
            resultBody.AppendChild(importedNode);
        }

        // Step 7: Encrypt the resulting document with a password.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Password = "Secret123"
        };
        string outputPath = "extracted_encrypted.docx";
        resultDoc.Save(outputPath, saveOptions);

        // Step 8: Validate that the encrypted file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Encrypted output file was not created.");
    }
}
