using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create a sample document with footnotes and bookmarks that define the extraction range.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Paragraph before the range.");

        // Start bookmark – marks the beginning of the extraction range.
        builder.StartBookmark("StartRange");
        builder.Writeln("Start paragraph.");
        builder.EndBookmark("StartRange");

        // Paragraphs with footnotes inside the range.
        builder.Writeln("Paragraph with first footnote.");
        builder.InsertFootnote(FootnoteType.Footnote, "Content of footnote 1.");

        builder.Writeln("Paragraph with second footnote.");
        builder.InsertFootnote(FootnoteType.Footnote, "Content of footnote 2.");

        // End bookmark – marks the end of the extraction range.
        builder.StartBookmark("EndRange");
        builder.Writeln("End paragraph.");
        builder.EndBookmark("EndRange");

        builder.Writeln("Paragraph after the range.");

        // Save the sample document.
        const string inputPath = "footnote-input.docx";
        doc.Save(inputPath);

        // Load the document for extraction.
        Document loaded = new Document(inputPath);

        // Retrieve the start and end bookmarks.
        Bookmark startBookmark = loaded.Range.Bookmarks["StartRange"];
        Bookmark endBookmark = loaded.Range.Bookmarks["EndRange"];
        if (startBookmark == null || endBookmark == null)
            throw new InvalidOperationException("Required bookmarks were not found.");

        // Resolve the paragraphs that contain the bookmark starts.
        Paragraph startParagraph = startBookmark.BookmarkStart?.ParentNode as Paragraph;
        Paragraph endParagraph = endBookmark.BookmarkStart?.ParentNode as Paragraph;
        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Start or end paragraph could not be resolved.");

        // Both bookmarks reside in the same body.
        CompositeNode body = startParagraph.ParentNode as CompositeNode;
        if (body == null)
            throw new InvalidOperationException("Unable to locate the document body.");

        int startIndex = body.IndexOf(startParagraph);
        int endIndex = body.IndexOf(endParagraph);
        if (startIndex < 0 || endIndex < 0)
            throw new InvalidOperationException("Start or end paragraph not found in body.");

        // Extract footnotes whose anchor paragraph lies within the specified range.
        int extractedCount = 0;
        foreach (Footnote footnote in loaded.GetChildNodes(NodeType.Footnote, true))
        {
            Paragraph anchorParagraph = footnote.ParentParagraph;
            if (anchorParagraph == null)
                continue;

            // Ensure the anchor paragraph belongs to the same body.
            CompositeNode anchorBody = anchorParagraph.ParentNode as CompositeNode;
            if (anchorBody != body)
                continue;

            int anchorIndex = body.IndexOf(anchorParagraph);
            if (anchorIndex >= startIndex && anchorIndex <= endIndex)
            {
                string fileName = $"footnote-{extractedCount}.txt";
                File.WriteAllText(fileName, footnote.GetText().Trim());
                extractedCount++;
            }
        }

        if (extractedCount == 0)
            throw new InvalidOperationException("No footnote files were generated.");

        // Optional: indicate success.
        Console.WriteLine($"Successfully extracted {extractedCount} footnote(s).");
    }
}
