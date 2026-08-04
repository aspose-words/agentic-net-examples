using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // 1. Create a sample document with several paragraphs.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");
        builder.Writeln("Paragraph 3");
        builder.Writeln("Paragraph 4");
        sourceDoc.Save("source.docx");

        // 2. Load the document for extraction.
        Document loadedDoc = new Document("source.docx");

        // 3. Define start and end paragraphs (intentionally reversed).
        Paragraph startParagraph = loadedDoc.FirstSection.Body.Paragraphs[2]; // "Paragraph 3"
        Paragraph endParagraph   = loadedDoc.FirstSection.Body.Paragraphs[1]; // "Paragraph 2"

        // 4. Validate node order and swap if necessary.
        int startIndex = loadedDoc.FirstSection.Body.Paragraphs.IndexOf(startParagraph);
        int endIndex   = loadedDoc.FirstSection.Body.Paragraphs.IndexOf(endParagraph);

        if (startIndex > endIndex)
        {
            Console.WriteLine(
                $"Warning: start paragraph index ({startIndex}) is after end paragraph index ({endIndex}). " +
                "Swapping the boundaries to continue extraction.");

            int temp = startIndex;
            startIndex = endIndex;
            endIndex = temp;
        }

        // 5. Prepare the result document (empty structure).
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren();

        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);

        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // 6. Import and append each paragraph from start to end (inclusive).
        NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);

        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph srcPara = loadedDoc.FirstSection.Body.Paragraphs[i];
            Node importedNode = importer.ImportNode(srcPara, true);
            resultBody.AppendChild(importedNode);
        }

        // 7. Save the extracted content.
        string outputPath = "extracted.docx";
        resultDoc.Save(outputPath);

        // 8. Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The extracted document was not created.");

        Console.WriteLine("Extraction completed successfully.");
    }
}
