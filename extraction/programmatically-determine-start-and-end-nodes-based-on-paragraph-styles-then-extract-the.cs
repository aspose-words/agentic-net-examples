using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample document with styled paragraphs.
        const string sourcePath = "source.docx";
        CreateSampleDocument(sourcePath);

        // Load the document.
        Document sourceDoc = new Document(sourcePath);

        // Define the styles that mark the start and end of the segment to extract.
        const string startStyleName = "Heading 1";
        const string endStyleName = "Heading 2";

        // Locate the start and end paragraphs based on their styles.
        Paragraph startParagraph = null;
        Paragraph endParagraph = null;
        int startIndex = -1;
        int endIndex = -1;

        NodeCollection paragraphNodes = sourceDoc.FirstSection.Body.Paragraphs;
        for (int i = 0; i < paragraphNodes.Count; i++)
        {
            Paragraph para = (Paragraph)paragraphNodes[i];
            string style = para.ParagraphFormat.StyleName;

            if (startParagraph == null && style == startStyleName)
            {
                startParagraph = para;
                startIndex = i;
            }
            else if (startParagraph != null && style == endStyleName)
            {
                endParagraph = para;
                endIndex = i;
                break;
            }
        }

        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Styled start or end paragraph not found.");

        // Build a new document that will contain the extracted segment.
        Document extractedDoc = new Document();
        extractedDoc.RemoveAllChildren();

        // Create a fresh section and body for the destination document.
        Section section = new Section(extractedDoc);
        extractedDoc.AppendChild(section);

        Body body = new Body(extractedDoc);
        section.AppendChild(body);

        // Import and copy all paragraphs from start to end (inclusive) into the new document.
        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph srcPara = (Paragraph)paragraphNodes[i];
            // Import the node so that it belongs to the destination document.
            Node importedNode = extractedDoc.ImportNode(srcPara, true);
            body.AppendChild(importedNode);
        }

        // Save the extracted content.
        const string resultPath = "extracted.docx";
        extractedDoc.Save(resultPath);

        // Verify that the output file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The extracted document was not saved.");
    }

    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Paragraph that marks the start (Heading 1).
        builder.ParagraphFormat.StyleName = "Heading 1";
        builder.Writeln("Start of extracted segment");

        // Some normal paragraphs inside the segment.
        builder.ParagraphFormat.StyleName = "Normal";
        builder.Writeln("First paragraph inside segment.");
        builder.Writeln("Second paragraph inside segment.");

        // Paragraph that marks the end (Heading 2).
        builder.ParagraphFormat.StyleName = "Heading 2";
        builder.Writeln("End of extracted segment");

        // Additional content after the segment.
        builder.ParagraphFormat.StyleName = "Normal";
        builder.Writeln("Paragraph after the extracted segment.");

        doc.Save(filePath);
    }
}
