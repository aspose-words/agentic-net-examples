using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample source document with several paragraphs.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");
        builder.Writeln("Paragraph 3");
        builder.Writeln("Paragraph 4");

        // Identify the start and end nodes for extraction (Paragraph 1 and Paragraph 3).
        Paragraph startParagraph = (Paragraph)sourceDoc.GetChild(NodeType.Paragraph, 0, true);
        Paragraph endParagraph = (Paragraph)sourceDoc.GetChild(NodeType.Paragraph, 2, true);

        // Prepare a new document that will contain the extracted content.
        Document extractedDoc = new Document();
        extractedDoc.EnsureMinimum(); // Guarantees at least one section/body/paragraph.

        // Use NodeImporter to import nodes while preserving source formatting.
        NodeImporter importer = new NodeImporter(sourceDoc, extractedDoc, ImportFormatMode.KeepSourceFormatting);

        // Get all paragraphs from the source body.
        var allParagraphs = sourceDoc.FirstSection.Body.GetChildNodes(NodeType.Paragraph, true)
                                 .Cast<Paragraph>()
                                 .ToList();

        // Determine indices of the start and end paragraphs.
        int startIndex = allParagraphs.IndexOf(startParagraph);
        int endIndex = allParagraphs.IndexOf(endParagraph);

        if (startIndex == -1 || endIndex == -1 || startIndex > endIndex)
            throw new InvalidOperationException("Invalid start or end node for extraction.");

        // Import each paragraph within the range into the new document.
        foreach (var para in allParagraphs.Skip(startIndex).Take(endIndex - startIndex + 1))
        {
            Node importedNode = importer.ImportNode(para, true);
            extractedDoc.FirstSection.Body.AppendChild(importedNode);
        }

        // Encrypt the extracted document with a password.
        string outputPath = "ExtractedEncrypted.docx";
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.Password = "Secret123";
        extractedDoc.Save(outputPath, saveOptions);

        // Verify that the file was created and is encrypted.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The encrypted document was not saved.", outputPath);

        var fileInfo = FileFormatUtil.DetectFileFormat(outputPath);
        if (!fileInfo.IsEncrypted)
            throw new InvalidOperationException("The document is not encrypted as expected.");

        // Optional: indicate success (no interactive prompts required).
        Console.WriteLine("Extraction and encryption completed successfully.");
    }
}
