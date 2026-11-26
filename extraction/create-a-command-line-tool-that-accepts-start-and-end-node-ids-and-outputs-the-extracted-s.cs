using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main(string[] args)
    {
        // 1. Create a sample source document (sample.docx) with several paragraphs.
        const string sourcePath = "sample.docx";
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");
        builder.Writeln("Paragraph 3");
        builder.Writeln("Paragraph 4");
        sourceDoc.Save(sourcePath);

        // 2. Load the document that will be used for extraction.
        Document loadedDoc = new Document(sourcePath);

        // 3. Determine start and end paragraph indices.
        int startIndex;
        int endIndex;

        if (args.Length >= 2 &&
            int.TryParse(args[0], out int parsedStart) &&
            int.TryParse(args[1], out int parsedEnd))
        {
            startIndex = parsedStart;
            endIndex = parsedEnd;
        }
        else
        {
            startIndex = 1; // second paragraph (zero‑based)
            endIndex = 2;   // third paragraph
        }

        // Validate indices.
        ParagraphCollection paragraphs = loadedDoc.FirstSection.Body.Paragraphs;
        if (startIndex < 0 || endIndex < 0 ||
            startIndex >= paragraphs.Count || endIndex >= paragraphs.Count ||
            startIndex > endIndex)
        {
            throw new InvalidOperationException("Invalid start or end paragraph indexes.");
        }

        // 4. Locate the start and end nodes.
        Node startNode = paragraphs[startIndex];
        Node endNode = paragraphs[endIndex];

        // 5. Build a new document that will contain the extracted range.
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren();

        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);

        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // 6. Import and copy nodes from startNode to endNode (inclusive).
        NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);
        for (Node cur = startNode; cur != null; cur = cur.NextSibling)
        {
            if (cur.NodeType == NodeType.Paragraph || cur.NodeType == NodeType.Table)
            {
                Node importedNode = importer.ImportNode(cur, true);
                resultBody.AppendChild(importedNode);
            }

            if (cur == endNode)
                break;
        }

        // 7. Save the extracted segment as PDF.
        const string outputPdf = "extracted.pdf";
        resultDoc.Save(outputPdf, SaveFormat.Pdf);

        // 8. Verify that the PDF file was created.
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException("The PDF output was not created.");

        Console.WriteLine($"Extraction completed. PDF saved to '{outputPdf}'.");
    }
}
