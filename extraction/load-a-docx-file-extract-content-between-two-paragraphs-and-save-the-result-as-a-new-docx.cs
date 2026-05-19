using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample source document with several paragraphs.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2"); // Start marker
        builder.Writeln("Paragraph 3"); // End marker
        builder.Writeln("Paragraph 4");

        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the document from the file system.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // -----------------------------------------------------------------
        // 3. Locate the start and end paragraphs (by index in this example).
        // -----------------------------------------------------------------
        Paragraph startParagraph = loadedDoc.FirstSection.Body.Paragraphs[1];
        Paragraph endParagraph = loadedDoc.FirstSection.Body.Paragraphs[2];

        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Required paragraphs were not found.");

        // -----------------------------------------------------------------
        // 4. Validate the range.
        // -----------------------------------------------------------------
        int startIndex = loadedDoc.FirstSection.Body.Paragraphs.IndexOf(startParagraph);
        int endIndex = loadedDoc.FirstSection.Body.Paragraphs.IndexOf(endParagraph);

        if (startIndex < 0 || endIndex < 0 || endIndex < startIndex)
            throw new InvalidOperationException("Invalid paragraph range.");

        // -----------------------------------------------------------------
        // 5. Prepare the destination document.
        // -----------------------------------------------------------------
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren(); // Clear the default nodes.

        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);

        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // -----------------------------------------------------------------
        // 6. Import the selected paragraphs into the new document.
        //    Use NodeImporter to keep source formatting and avoid cross‑document
        //    ownership issues.
        // -----------------------------------------------------------------
        NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);

        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph srcParagraph = loadedDoc.FirstSection.Body.Paragraphs[i];
            // Import the paragraph (deep clone) into the destination document.
            Node importedNode = importer.ImportNode(srcParagraph, true);
            resultBody.AppendChild(importedNode);
        }

        // -----------------------------------------------------------------
        // 7. Save the extracted content as a new DOCX file.
        // -----------------------------------------------------------------
        const string resultPath = "extracted.docx";
        resultDoc.Save(resultPath);

        // -----------------------------------------------------------------
        // 8. Verify that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The extracted document was not created.");

        // -----------------------------------------------------------------
        // 9. Indicate success.
        // -----------------------------------------------------------------
        Console.WriteLine("Extraction completed successfully.");
    }
}
