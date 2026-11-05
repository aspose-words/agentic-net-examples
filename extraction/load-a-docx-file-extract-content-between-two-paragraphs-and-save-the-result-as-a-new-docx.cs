using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -------------------------------------------------
        // 1. Create a sample source document with several paragraphs.
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");
        builder.Writeln("Paragraph 3");
        builder.Writeln("Paragraph 4");

        string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // -------------------------------------------------
        // 2. Load the document from the file system.
        // -------------------------------------------------
        Document loadedDoc = new Document(sourcePath);
        Body sourceBody = loadedDoc.FirstSection.Body;

        // -------------------------------------------------
        // 3. Identify the start and end paragraphs (inclusive).
        //    For this example we extract "Paragraph 2" through "Paragraph 3".
        // -------------------------------------------------
        Paragraph startParagraph = sourceBody.Paragraphs[1];
        Paragraph endParagraph = sourceBody.Paragraphs[2];

        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Boundary paragraphs not found.");

        int startIndex = sourceBody.Paragraphs.IndexOf(startParagraph);
        int endIndex = sourceBody.Paragraphs.IndexOf(endParagraph);

        if (startIndex < 0 || endIndex < 0 || endIndex < startIndex)
            throw new InvalidOperationException("Invalid paragraph boundaries.");

        // -------------------------------------------------
        // 4. Prepare the destination document.
        // -------------------------------------------------
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren();

        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);

        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // -------------------------------------------------
        // 5. Import (clone) each paragraph from the source range into the destination.
        //    Use NodeImporter to keep the nodes compatible with the new document.
        // -------------------------------------------------
        NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);

        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph para = sourceBody.Paragraphs[i];
            Node importedNode = importer.ImportNode(para, true);
            resultBody.AppendChild(importedNode);
        }

        // -------------------------------------------------
        // 6. Save the extracted content to a new DOCX file.
        // -------------------------------------------------
        string resultPath = "extracted.docx";
        resultDoc.Save(resultPath);

        // -------------------------------------------------
        // 7. Validate that the output file was created.
        // -------------------------------------------------
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The extracted document was not created.");
    }
}
