using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample source document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Paragraph before start.");
        builder.Writeln("Paragraph containing the start run:");
        // This run will be the start marker.
        builder.Font.Bold = true;
        builder.Write("StartRun");
        builder.Font.Bold = false;
        builder.Writeln(); // finish the paragraph.

        builder.Writeln("Middle paragraph 1.");
        builder.Writeln("Middle paragraph 2.");

        // End bookmark marker.
        builder.StartBookmark("EndMarker");
        builder.Writeln("Paragraph inside end bookmark.");
        builder.EndBookmark("EndMarker");

        builder.Writeln("Paragraph after end.");

        // Save the source document.
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // Load the document for processing.
        Document doc = new Document(sourcePath);

        // Locate the start Run node with the specific text.
        Run startRun = doc.GetChildNodes(NodeType.Run, true)
                         .OfType<Run>()
                         .FirstOrDefault(r => r.Text == "StartRun");
        if (startRun == null)
            throw new InvalidOperationException("Start run node not found.");

        // Locate the end bookmark.
        Bookmark endBookmark = doc.Range.Bookmarks["EndMarker"];
        if (endBookmark == null)
            throw new InvalidOperationException("End bookmark not found.");

        // Determine the parent paragraphs of the start run and end bookmark.
        Paragraph startParagraph = startRun.GetAncestor(NodeType.Paragraph) as Paragraph;
        Paragraph endParagraph = endBookmark.BookmarkStart.GetAncestor(NodeType.Paragraph) as Paragraph;
        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Unable to determine boundary paragraphs.");

        // Find the indices of the boundary paragraphs within the body.
        Body body = doc.FirstSection.Body;
        int startIndex = body.Paragraphs.IndexOf(startParagraph);
        int endIndex = body.Paragraphs.IndexOf(endParagraph);
        if (startIndex < 0 || endIndex < 0 || startIndex >= endIndex)
            throw new InvalidOperationException("Invalid paragraph boundaries for extraction.");

        // Create a new document to hold the extracted content.
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren();

        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);

        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // Clone (import) and append each paragraph between the boundaries (exclusive).
        for (int i = startIndex + 1; i < endIndex; i++)
        {
            Paragraph para = body.Paragraphs[i];
            Node imported = resultDoc.ImportNode(para, true, ImportFormatMode.KeepSourceFormatting);
            resultBody.AppendChild(imported);
        }

        // Validate that something was extracted.
        if (resultBody.Paragraphs.Count == 0)
            throw new InvalidOperationException("No content was extracted between the specified nodes.");

        // Save the extracted content.
        const string outputPath = "extracted.docx";
        resultDoc.Save(outputPath);

        // Verify output file creation.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Extraction output file was not created.");
    }
}
