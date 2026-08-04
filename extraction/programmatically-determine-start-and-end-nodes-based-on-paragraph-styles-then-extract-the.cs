using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample source document with styled paragraphs.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Paragraph with style "Heading 1" – start marker.
        builder.ParagraphFormat.StyleName = "Heading 1";
        builder.Writeln("Start of the extracted segment");

        // Normal paragraphs.
        builder.ParagraphFormat.StyleName = "Normal";
        builder.Writeln("Paragraph A");
        builder.Writeln("Paragraph B");

        // Paragraph with style "Heading 2" – end marker.
        builder.ParagraphFormat.StyleName = "Heading 2";
        builder.Writeln("End of the extracted segment");

        // Additional content after the end marker.
        builder.ParagraphFormat.StyleName = "Normal";
        builder.Writeln("Paragraph C");

        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the document for processing.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // Locate the start and end paragraphs based on their styles.
        Paragraph startParagraph = null;
        Paragraph endParagraph = null;

        foreach (Paragraph para in loadedDoc.FirstSection.Body.Paragraphs)
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

        if (startParagraph == null)
            throw new InvalidOperationException("Start paragraph with style 'Heading 1' not found.");
        if (endParagraph == null)
            throw new InvalidOperationException("End paragraph with style 'Heading 2' not found.");

        // -----------------------------------------------------------------
        // 3. Build a new document containing the extracted range.
        // -----------------------------------------------------------------
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren();

        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);

        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // Determine the indices of the start and end paragraphs within the body.
        NodeCollection bodyParagraphs = loadedDoc.FirstSection.Body.GetChildNodes(NodeType.Paragraph, true);
        int startIndex = bodyParagraphs.IndexOf(startParagraph);
        int endIndex = bodyParagraphs.IndexOf(endParagraph);

        if (startIndex < 0 || endIndex < 0 || endIndex < startIndex)
            throw new InvalidOperationException("Invalid paragraph indices for extraction.");

        // Use NodeImporter to copy nodes from the source document to the destination document.
        NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);

        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph srcPara = (Paragraph)bodyParagraphs[i];
            Node importedNode = importer.ImportNode(srcPara, true);
            resultBody.AppendChild(importedNode);
        }

        // Verify that the extracted document contains the expected number of paragraphs.
        int expectedCount = endIndex - startIndex + 1;
        if (resultDoc.FirstSection.Body.Paragraphs.Count != expectedCount)
            throw new InvalidOperationException("Paragraph count mismatch after extraction.");

        // -----------------------------------------------------------------
        // 4. Save the extracted segment.
        // -----------------------------------------------------------------
        const string resultPath = "extracted.docx";
        resultDoc.Save(resultPath);

        // Ensure the output file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The extracted document was not created.");
    }
}
