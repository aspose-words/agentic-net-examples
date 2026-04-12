using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample source document with three paragraphs.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");
        builder.Writeln("Paragraph 3");

        // Save the source document (optional, just to have a physical file).
        string sourcePath = Path.Combine(Directory.GetCurrentDirectory(), "Source.docx");
        sourceDoc.Save(sourcePath);

        // Retrieve the three paragraphs from the document body.
        Paragraph para1 = sourceDoc.FirstSection.Body.Paragraphs[0];
        Paragraph para2 = sourceDoc.FirstSection.Body.Paragraphs[1];
        Paragraph para3 = sourceDoc.FirstSection.Body.Paragraphs[2];

        // Attempt a correct extraction (para2 to para3).
        try
        {
            Document extractedCorrect = ExtractParagraphRange(sourceDoc, para2, para3);
            string correctPath = Path.Combine(Directory.GetCurrentDirectory(), "Extracted_Correct.docx");
            extractedCorrect.Save(correctPath);
            Console.WriteLine($"Correct extraction succeeded. File saved to: {correctPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Correct extraction failed: {ex.Message}");
        }

        // Attempt an incorrect extraction where the start node appears after the end node (para3 to para2).
        try
        {
            Document extractedInvalid = ExtractParagraphRange(sourceDoc, para3, para2);
            string invalidPath = Path.Combine(Directory.GetCurrentDirectory(), "Extracted_Invalid.docx");
            extractedInvalid.Save(invalidPath);
            Console.WriteLine($"Invalid extraction succeeded (unexpected). File saved to: {invalidPath}");
        }
        catch (InvalidOperationException ex)
        {
            // Expected path for reversed nodes.
            Console.WriteLine($"Invalid extraction correctly detected error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Any other unexpected errors.
            Console.WriteLine($"Unexpected error during invalid extraction: {ex.Message}");
        }
    }

    /// <summary>
    /// Extracts a range of consecutive paragraphs from <paramref name="doc"/> inclusive of <paramref name="startParagraph"/>
    /// and <paramref name="endParagraph"/>. Throws an InvalidOperationException if the start paragraph appears after the end paragraph.
    /// </summary>
    private static Document ExtractParagraphRange(Document doc, Paragraph startParagraph, Paragraph endParagraph)
    {
        if (doc == null) throw new ArgumentNullException(nameof(doc));
        if (startParagraph == null) throw new ArgumentNullException(nameof(startParagraph));
        if (endParagraph == null) throw new ArgumentNullException(nameof(endParagraph));

        // Ensure both paragraphs belong to the same document.
        if (!ReferenceEquals(startParagraph.Document, doc) || !ReferenceEquals(endParagraph.Document, doc))
            throw new ArgumentException("Paragraphs must belong to the provided document.");

        // Determine the indices of the start and end paragraphs within the body.
        Body body = doc.FirstSection.Body;
        int startIndex = body.Paragraphs.IndexOf(startParagraph);
        int endIndex = body.Paragraphs.IndexOf(endParagraph);

        if (startIndex == -1 || endIndex == -1)
            throw new InvalidOperationException("One or both paragraphs were not found in the document body.");

        // Validate ordering.
        if (startIndex > endIndex)
            throw new InvalidOperationException("The start paragraph appears after the end paragraph. Extraction aborted.");

        // Create a new empty document to hold the extracted content.
        Document result = new Document();
        result.RemoveAllChildren(); // Remove the default section/paragraph.
        Section newSection = new Section(result);
        result.AppendChild(newSection);
        Body newBody = new Body(result);
        newSection.AppendChild(newBody);

        // Import each paragraph in the range and append to the new body.
        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph srcParagraph = (Paragraph)body.Paragraphs[i];
            // Use ImportNode to preserve formatting.
            Paragraph importedParagraph = (Paragraph)result.ImportNode(srcParagraph, true);
            newBody.AppendChild(importedParagraph);
        }

        // Validate that at least one paragraph was extracted.
        if (newBody.Paragraphs.Count == 0)
            throw new InvalidOperationException("No paragraphs were extracted; the resulting document is empty.");

        return result;
    }
}
