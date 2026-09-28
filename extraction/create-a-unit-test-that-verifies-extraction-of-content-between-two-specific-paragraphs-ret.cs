using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // ------------------------------------------------------------
        // 1. Create a source document with distinct paragraph styling.
        // ------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Paragraph 0 - normal");

        builder.Font.Bold = true;
        builder.Writeln("Paragraph 1 - bold");

        builder.Font.Bold = false;
        builder.Font.Italic = true;
        builder.Writeln("Paragraph 2 - italic");

        builder.Font.Italic = false;
        builder.Writeln("Paragraph 3 - normal");

        // Save the source document for manual inspection (optional).
        sourceDoc.Save("source.docx");

        // ------------------------------------------------------------
        // 2. Identify the start and end paragraphs (indexes 1 and 2).
        // ------------------------------------------------------------
        Paragraph startPara = sourceDoc.FirstSection.Body.Paragraphs[1];
        Paragraph endPara = sourceDoc.FirstSection.Body.Paragraphs[2];

        if (startPara == null || endPara == null)
            throw new InvalidOperationException("Boundary paragraphs not found.");

        // ------------------------------------------------------------
        // 3. Create a new document that will hold the extracted content.
        // ------------------------------------------------------------
        Document extractedDoc = new Document();
        extractedDoc.RemoveAllChildren(); // Remove the default empty section.

        // Build a clean document structure: Section -> Body.
        Section newSection = new Section(extractedDoc);
        extractedDoc.AppendChild(newSection);
        Body newBody = new Body(extractedDoc);
        newSection.AppendChild(newBody);

        // ------------------------------------------------------------
        // 4. Import (clone) the selected paragraphs into the new document,
        //    preserving all formatting.
        // ------------------------------------------------------------
        Paragraph importedStart = (Paragraph)extractedDoc.ImportNode(startPara, true, ImportFormatMode.KeepSourceFormatting);
        Paragraph importedEnd = (Paragraph)extractedDoc.ImportNode(endPara, true, ImportFormatMode.KeepSourceFormatting);

        newBody.AppendChild(importedStart);
        newBody.AppendChild(importedEnd);

        // ------------------------------------------------------------
        // 5. Validate that the imported paragraphs retain the original styling.
        // ------------------------------------------------------------
        Paragraph clonedStart = extractedDoc.FirstSection.Body.Paragraphs[0];
        Paragraph clonedEnd = extractedDoc.FirstSection.Body.Paragraphs[1];

        if (clonedStart == null || clonedEnd == null)
            throw new InvalidOperationException("Cloned paragraphs were not added correctly.");

        // Each paragraph contains a single run because we used Writeln.
        Run sourceStartRun = startPara.Runs[0];
        Run sourceEndRun = endPara.Runs[0];
        Run clonedStartRun = clonedStart.Runs[0];
        Run clonedEndRun = clonedEnd.Runs[0];

        if (sourceStartRun == null || sourceEndRun == null || clonedStartRun == null || clonedEndRun == null)
            throw new InvalidOperationException("Expected runs were not found.");

        // Compare formatting properties.
        if (clonedStartRun.Font.Bold != sourceStartRun.Font.Bold ||
            clonedStartRun.Font.Italic != sourceStartRun.Font.Italic)
            throw new InvalidOperationException("Styling of the first extracted paragraph does not match the source.");

        if (clonedEndRun.Font.Bold != sourceEndRun.Font.Bold ||
            clonedEndRun.Font.Italic != sourceEndRun.Font.Italic)
            throw new InvalidOperationException("Styling of the second extracted paragraph does not match the source.");

        // ------------------------------------------------------------
        // 6. Save the extracted document.
        // ------------------------------------------------------------
        extractedDoc.Save("extracted.docx");

        // Indicate success.
        Console.WriteLine("Extraction and styling verification succeeded.");
    }
}
