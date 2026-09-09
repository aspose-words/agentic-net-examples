using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create a sample document with bookmarks and footnotes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start of the range.
        builder.StartBookmark("StartRange");
        builder.Writeln("Paragraph inside range with footnote 1.");
        builder.InsertFootnote(FootnoteType.Footnote, "Footnote 1 content.");
        builder.Writeln("Paragraph inside range without footnote.");
        builder.EndBookmark("StartRange");

        // Paragraph outside the range (should not be extracted).
        builder.Writeln("Paragraph outside range with footnote 2.");
        builder.InsertFootnote(FootnoteType.Footnote, "Footnote 2 content.");

        // End of the range.
        builder.StartBookmark("EndRange");
        builder.Writeln("Paragraph inside range with footnote 3.");
        builder.InsertFootnote(FootnoteType.Footnote, "Footnote 3 content.");
        builder.EndBookmark("EndRange");

        // Save the sample document.
        const string inputPath = "sample.docx";
        doc.Save(inputPath);

        // Load the document for extraction.
        Document loaded = new Document(inputPath);

        // Retrieve the start and end bookmarks.
        Bookmark startBookmark = loaded.Range.Bookmarks["StartRange"];
        Bookmark endBookmark = loaded.Range.Bookmarks["EndRange"];
        if (startBookmark == null || endBookmark == null)
            throw new InvalidOperationException("Required bookmarks were not found.");

        // Determine the paragraphs that bound the range.
        Paragraph startParagraph = startBookmark.BookmarkStart.ParentNode as Paragraph;
        Paragraph endParagraph = endBookmark.BookmarkEnd.ParentNode as Paragraph;
        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Bookmark boundaries are not paragraphs.");

        // Collect all paragraphs in document order.
        NodeCollection allParagraphs = loaded.GetChildNodes(NodeType.Paragraph, true);
        int startIndex = -1;
        int endIndex = -1;
        for (int i = 0; i < allParagraphs.Count; i++)
        {
            if (allParagraphs[i] == startParagraph) startIndex = i;
            if (allParagraphs[i] == endParagraph) endIndex = i;
        }
        if (startIndex == -1 || endIndex == -1 || startIndex > endIndex)
            throw new InvalidOperationException("Invalid bookmark range.");

        // Extract footnotes that belong to paragraphs within the range.
        int footnoteCounter = 0;
        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph para = allParagraphs[i] as Paragraph;
            if (para == null) continue;

            NodeCollection footnotes = para.GetChildNodes(NodeType.Footnote, true);
            foreach (Footnote footnote in footnotes.OfType<Footnote>())
            {
                string fileName = $"footnote-{footnoteCounter}.txt";
                File.WriteAllText(fileName, footnote.GetText().Trim());
                footnoteCounter++;
            }
        }

        if (footnoteCounter == 0)
            throw new InvalidOperationException("No footnote files were generated.");
    }
}
