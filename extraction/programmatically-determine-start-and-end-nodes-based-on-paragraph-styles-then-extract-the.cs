using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample document with styled paragraphs.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Paragraph with style "Heading 1" – start marker.
        builder.ParagraphFormat.StyleName = "Heading 1";
        builder.Writeln("Start Section");

        // Normal paragraphs – content to be extracted.
        builder.ParagraphFormat.StyleName = "Normal";
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");

        // Paragraph with style "Heading 2" – end marker.
        builder.ParagraphFormat.StyleName = "Heading 2";
        builder.Writeln("End Section");

        // Save the source document locally.
        const string sourcePath = "styled-input.docx";
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the document for processing.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);
        ParagraphCollection paragraphs = loadedDoc.FirstSection.Body.Paragraphs;

        // Locate start and end paragraphs by their styles.
        Paragraph startParagraph = null;
        Paragraph endParagraph = null;

        foreach (Paragraph para in paragraphs)
        {
            string styleName = para.ParagraphFormat.StyleName;

            if (startParagraph == null && styleName == "Heading 1")
                startParagraph = para;

            if (endParagraph == null && styleName == "Heading 2")
                endParagraph = para;

            if (startParagraph != null && endParagraph != null)
                break;
        }

        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Start or end styled paragraph not found.");

        int startIndex = paragraphs.IndexOf(startParagraph);
        int endIndex = paragraphs.IndexOf(endParagraph);

        if (startIndex > endIndex)
            throw new InvalidOperationException("Start paragraph occurs after end paragraph.");

        // -----------------------------------------------------------------
        // 3. Build a new document that will contain the extracted range.
        // -----------------------------------------------------------------
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren();

        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);

        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // Use NodeImporter to copy nodes from the source document to the result document.
        NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);

        for (int i = startIndex; i <= endIndex; i++)
        {
            Node importedNode = importer.ImportNode(paragraphs[i], true);
            resultBody.AppendChild(importedNode);
        }

        // Validate that the expected number of paragraphs were copied.
        int expectedCount = endIndex - startIndex + 1;
        if (resultBody.Paragraphs.Count != expectedCount)
            throw new InvalidOperationException("Extracted paragraph count mismatch.");

        // -----------------------------------------------------------------
        // 4. Save the extracted content.
        // -----------------------------------------------------------------
        const string resultPath = "extracted-styled.docx";
        resultDoc.Save(resultPath);

        // Verify that the output file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("Extraction output file was not created.");
    }
}
