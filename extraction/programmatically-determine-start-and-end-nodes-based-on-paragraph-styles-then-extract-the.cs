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

        // Heading 1 - start marker (bold).
        builder.Font.Bold = true;
        builder.ParagraphFormat.StyleName = "Heading 1";
        builder.Writeln("Start Section");

        // Normal paragraphs - content.
        builder.Font.Bold = false;
        builder.ParagraphFormat.StyleName = "Normal";
        builder.Writeln("First content paragraph.");
        builder.Writeln("Second content paragraph.");

        // Heading 2 - end marker (italic).
        builder.Font.Italic = true;
        builder.ParagraphFormat.StyleName = "Heading 2";
        builder.Writeln("End Section");

        // Reset formatting for any further text.
        builder.Font.Italic = false;
        builder.Font.Bold = false;
        builder.ParagraphFormat.StyleName = "Normal";

        // Save the source document.
        const string sourcePath = "styled-input.docx";
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the document and locate start/end paragraphs by style.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);
        Paragraph startParagraph = null;
        Paragraph endParagraph = null;

        NodeCollection paragraphs = loadedDoc.FirstSection.Body.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            string styleName = para.ParagraphFormat.StyleName;
            if (styleName == "Heading 1" && startParagraph == null)
                startParagraph = para;
            else if (styleName == "Heading 2" && endParagraph == null)
                endParagraph = para;
        }

        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Required styled paragraphs were not found.");

        // Determine the indices of the start and end paragraphs.
        int startIndex = paragraphs.IndexOf(startParagraph);
        int endIndex = paragraphs.IndexOf(endParagraph);
        if (startIndex > endIndex)
            throw new InvalidOperationException("Start paragraph occurs after end paragraph.");

        // -----------------------------------------------------------------
        // 3. Extract the range of paragraphs (inclusive) into a new document.
        // -----------------------------------------------------------------
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren();

        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);

        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // Import each paragraph from the source document into the result document.
        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph srcPara = (Paragraph)paragraphs[i];
            // ImportNode creates a copy that belongs to resultDoc.
            Node importedNode = resultDoc.ImportNode(srcPara, true);
            resultBody.AppendChild(importedNode);
        }

        // Validate that styling was preserved (first paragraph should be bold).
        Paragraph clonedFirst = resultDoc.FirstSection.Body.Paragraphs[0];
        if (clonedFirst.Runs.Count == 0 || !clonedFirst.Runs[0].Font.Bold)
            throw new InvalidOperationException("Extracted paragraph styling does not match the source.");

        // -----------------------------------------------------------------
        // 4. Save the extracted segment.
        // -----------------------------------------------------------------
        const string resultPath = "styled-extraction.docx";
        resultDoc.Save(resultPath);

        // Verify that the output file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("Expected extraction output was not created.");
    }
}
