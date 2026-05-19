using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -------------------------------------------------
        // 1. Create a sample source document with a start run and an end bookmark.
        // -------------------------------------------------
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);

        builder.Writeln("Paragraph before start.");
        // Distinct run that will be used as the start node.
        builder.Write("StartRun");
        builder.Writeln(); // End the paragraph.

        builder.Writeln("Middle paragraph 1.");
        builder.Writeln("Middle paragraph 2.");

        // End bookmark.
        builder.StartBookmark("EndBookmark");
        builder.Writeln("Content inside end bookmark.");
        builder.EndBookmark("EndBookmark");

        builder.Writeln("Paragraph after end.");

        const string sourcePath = "source.docx";
        source.Save(sourcePath);

        // -------------------------------------------------
        // 2. Load the document for processing.
        // -------------------------------------------------
        Document loaded = new Document(sourcePath);

        // -------------------------------------------------
        // 3. Locate the start Run node with the exact text "StartRun".
        // -------------------------------------------------
        Run startRun = null;
        foreach (Run run in loaded.GetChildNodes(NodeType.Run, true))
        {
            if (run.Text == "StartRun")
            {
                startRun = run;
                break;
            }
        }

        if (startRun == null)
            throw new InvalidOperationException("Start run node not found.");

        // -------------------------------------------------
        // 4. Locate the end bookmark.
        // -------------------------------------------------
        Bookmark endBookmark = loaded.Range.Bookmarks["EndBookmark"];
        if (endBookmark == null)
            throw new InvalidOperationException("End bookmark not found.");

        // -------------------------------------------------
        // 5. Determine the bounding paragraphs.
        // -------------------------------------------------
        Paragraph startParagraph = (Paragraph)startRun.GetAncestor(NodeType.Paragraph);
        Paragraph endParagraph = (Paragraph)endBookmark.BookmarkStart.GetAncestor(NodeType.Paragraph);

        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Unable to determine bounding paragraphs.");

        Body sourceBody = startParagraph.ParentNode as Body;
        if (sourceBody == null)
            throw new InvalidOperationException("Source body not found.");

        // Use the Paragraphs collection to obtain indices.
        int startIndex = sourceBody.Paragraphs.IndexOf(startParagraph);
        int endIndex = sourceBody.Paragraphs.IndexOf(endParagraph);

        if (startIndex < 0 || endIndex < 0 || startIndex > endIndex)
            throw new InvalidOperationException("Invalid paragraph indices for extraction.");

        // -------------------------------------------------
        // 6. Prepare the destination document (empty).
        // -------------------------------------------------
        Document destination = new Document();
        destination.RemoveAllChildren(); // Remove the default empty section/paragraph.
        Section destSection = new Section(destination);
        destination.AppendChild(destSection);
        Body destBody = new Body(destination);
        destSection.AppendChild(destBody);

        // -------------------------------------------------
        // 7. Import the selected paragraphs into the destination document.
        // -------------------------------------------------
        NodeImporter importer = new NodeImporter(loaded, destination, ImportFormatMode.KeepSourceFormatting);
        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph para = sourceBody.Paragraphs[i];
            Node imported = importer.ImportNode(para, true);
            destBody.AppendChild(imported);
        }

        // -------------------------------------------------
        // 8. Save the extracted content.
        // -------------------------------------------------
        const string resultPath = "extracted.docx";
        destination.Save(resultPath);

        // -------------------------------------------------
        // 9. Validate that the output file was created.
        // -------------------------------------------------
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("Extraction failed: output file was not created.");
    }
}
