using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -------------------------------------------------
        // Create a sample source document with styled markers.
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Introduction paragraph.");

        // Start marker – style "Heading 1".
        builder.ParagraphFormat.StyleName = "Heading 1";
        builder.Writeln("Start of Styled Segment");

        // Content inside the segment with distinct formatting.
        builder.Font.Bold = true;
        builder.Writeln("Bold paragraph inside segment.");
        builder.Font.Bold = false;
        builder.Font.Italic = true;
        builder.Writeln("Italic paragraph inside segment.");
        builder.Font.Italic = false;

        // End marker – style "Heading 2".
        builder.ParagraphFormat.StyleName = "Heading 2";
        builder.Writeln("End of Styled Segment");

        // Normal paragraph after the segment.
        builder.ParagraphFormat.StyleName = "Normal";
        builder.Writeln("Conclusion paragraph.");

        // Save the source document locally.
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // -------------------------------------------------
        // Load the document for processing.
        // -------------------------------------------------
        Document loadedDoc = new Document(sourcePath);
        ParagraphCollection paragraphs = loadedDoc.FirstSection.Body.Paragraphs;

        // Locate the start and end paragraphs by their styles.
        Paragraph startParagraph = null;
        Paragraph endParagraph = null;

        foreach (Paragraph para in paragraphs)
        {
            string styleName = para.ParagraphFormat.StyleName;
            if (startParagraph == null && styleName == "Heading 1")
                startParagraph = para;
            else if (startParagraph != null && styleName == "Heading 2")
            {
                endParagraph = para;
                break;
            }
        }

        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Could not locate the start or end styled paragraphs.");

        int startIndex = paragraphs.IndexOf(startParagraph);
        int endIndex = paragraphs.IndexOf(endParagraph);

        if (startIndex < 0 || endIndex < 0 || endIndex < startIndex)
            throw new InvalidOperationException("Invalid paragraph indices for extraction.");

        // -------------------------------------------------
        // Build a new document containing the extracted range.
        // -------------------------------------------------
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren(); // Ensure a clean document structure.

        // Create a new section and body for the result document.
        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);
        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // Import each paragraph from the source range into the result document.
        for (int i = startIndex; i <= endIndex; i++)
        {
            // ImportNode clones the node and reassigns it to the destination document.
            Node importedNode = resultDoc.ImportNode(paragraphs[i], true);
            resultBody.AppendChild(importedNode);
        }

        // Save the extracted segment.
        const string resultPath = "extracted.docx";
        resultDoc.Save(resultPath);

        // Validate that the output file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The extracted document was not created.");
    }
}
