using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // 1. Create a sample source document with identifiable paragraphs.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Intro paragraph.");
        builder.Writeln("Start extraction paragraph.");   // index 1
        builder.Writeln("Paragraph inside extraction."); // index 2
        builder.Writeln("End extraction paragraph.");    // index 3
        builder.Writeln("After extraction paragraph.");
        source.Save("source.docx");

        // 2. Load the source document.
        Document loaded = new Document("source.docx");

        // 3. Locate the start and end paragraphs that bound the extraction range.
        Paragraph startPara = loaded.FirstSection.Body.Paragraphs[1]; // "Start extraction paragraph."
        Paragraph endPara   = loaded.FirstSection.Body.Paragraphs[3]; // "End extraction paragraph."

        if (startPara == null || endPara == null)
            throw new InvalidOperationException("Required boundary paragraphs were not found.");

        int startIndex = loaded.FirstSection.Body.Paragraphs.IndexOf(startPara);
        int endIndex   = loaded.FirstSection.Body.Paragraphs.IndexOf(endPara);

        if (startIndex < 0 || endIndex < 0 || endIndex < startIndex)
            throw new InvalidOperationException("Invalid paragraph range for extraction.");

        // 4. Prepare the destination document that will hold the extracted content.
        Document result = new Document();
        result.RemoveAllChildren();                     // clear default nodes
        Section resultSection = new Section(result);
        result.AppendChild(resultSection);
        Body resultBody = new Body(result);
        resultSection.AppendChild(resultBody);

        // 5. Import each paragraph from the source range into the destination.
        NodeImporter importer = new NodeImporter(loaded, result, ImportFormatMode.KeepSourceFormatting);

        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph srcParagraph = loaded.FirstSection.Body.Paragraphs[i];
            Node importedNode = importer.ImportNode(srcParagraph, true);
            resultBody.AppendChild(importedNode);
        }

        // 6. Encrypt the extracted document with a password and save it.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            Password = "Secret123"
        };
        string outputPath = "extracted_encrypted.docx";
        result.Save(outputPath, saveOptions);

        // 7. Verify that the encrypted file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The encrypted output file was not created.");
    }
}
