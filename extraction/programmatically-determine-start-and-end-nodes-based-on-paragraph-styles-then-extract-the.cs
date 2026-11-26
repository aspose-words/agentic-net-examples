using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // -------------------------------------------------
        // 1. Create a sample source document with custom styles.
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Add custom paragraph styles.
        Style startStyle = sourceDoc.Styles.Add(StyleType.Paragraph, "StartStyle");
        startStyle.Font.Bold = true;
        Style endStyle = sourceDoc.Styles.Add(StyleType.Paragraph, "EndStyle");
        endStyle.Font.Italic = true;

        // Paragraphs with the defined styles and normal paragraphs.
        builder.ParagraphFormat.StyleName = "StartStyle";
        builder.Writeln("This is the start styled paragraph.");

        builder.ParagraphFormat.StyleName = "Normal";
        builder.Writeln("First normal paragraph.");
        builder.Writeln("Second normal paragraph.");

        builder.ParagraphFormat.StyleName = "EndStyle";
        builder.Writeln("This is the end styled paragraph.");

        builder.ParagraphFormat.StyleName = "Normal";
        builder.Writeln("Paragraph after end.");

        // Save the source document locally.
        const string sourcePath = "styled-input.docx";
        sourceDoc.Save(sourcePath);

        // -------------------------------------------------
        // 2. Load the document and locate the start/end paragraphs.
        // -------------------------------------------------
        Document loadedDoc = new Document(sourcePath);
        Body sourceBody = loadedDoc.FirstSection.Body;

        Paragraph startParagraph = null;
        Paragraph endParagraph = null;

        foreach (Paragraph para in sourceBody.Paragraphs)
        {
            string styleName = para.ParagraphFormat.StyleName;
            if (styleName == "StartStyle" && startParagraph == null)
                startParagraph = para;
            else if (styleName == "EndStyle" && startParagraph != null && endParagraph == null)
                endParagraph = para;
        }

        if (startParagraph == null)
            throw new InvalidOperationException("Start styled paragraph not found.");
        if (endParagraph == null)
            throw new InvalidOperationException("End styled paragraph not found.");

        // Determine the indices of the boundary paragraphs.
        int startIndex = sourceBody.Paragraphs.IndexOf(startParagraph);
        int endIndex = sourceBody.Paragraphs.IndexOf(endParagraph);
        if (startIndex < 0 || endIndex < 0 || endIndex < startIndex)
            throw new InvalidOperationException("Invalid paragraph boundaries.");

        // -------------------------------------------------
        // 3. Build a new document that will contain the extracted segment.
        // -------------------------------------------------
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren();

        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);

        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // Use NodeImporter to import nodes from the source document into the result document.
        NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);

        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph srcPara = sourceBody.Paragraphs[i];
            Node importedNode = importer.ImportNode(srcPara, true);
            resultBody.AppendChild(importedNode);
        }

        // -------------------------------------------------
        // 4. Save the extracted segment and validate the output.
        // -------------------------------------------------
        const string resultPath = "extracted-styled.docx";
        resultDoc.Save(resultPath);

        if (!File.Exists(resultPath))
            throw new InvalidOperationException("Extraction output file was not created.");
    }
}
