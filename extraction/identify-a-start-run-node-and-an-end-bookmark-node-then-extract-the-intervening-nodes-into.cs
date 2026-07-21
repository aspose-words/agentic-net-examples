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
        // This run will be the start node we look for.
        builder.Write("StartRun");
        builder.Writeln(); // finish the paragraph.

        builder.Writeln("Middle paragraph 1.");
        builder.Writeln("Middle paragraph 2.");

        // End bookmark.
        builder.StartBookmark("EndBookmark");
        builder.Writeln("Paragraph after end.");
        builder.EndBookmark("EndBookmark");

        // Save the source document to a deterministic local file.
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

        // The bookmark start node marks the end of the extraction range.
        BookmarkStart endBookmarkStart = endBookmark.BookmarkStart;

        // -------------------------------------------------
        // 5. Determine the paragraphs that bound the extraction range.
        // -------------------------------------------------
        Paragraph startParagraph = startRun.ParentNode as Paragraph;
        Paragraph endParagraph = endBookmarkStart.ParentNode as Paragraph;

        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Unable to determine paragraph boundaries.");

        // -------------------------------------------------
        // 6. Prepare the destination document (empty body).
        // -------------------------------------------------
        Document result = new Document();
        result.RemoveAllChildren(); // Ensure a clean document.
        Section resultSection = new Section(result);
        result.AppendChild(resultSection);
        Body resultBody = new Body(result);
        resultSection.AppendChild(resultBody);

        // -------------------------------------------------
        // 7. Importer to copy nodes while preserving formatting.
        // -------------------------------------------------
        NodeImporter importer = new NodeImporter(loaded, result, ImportFormatMode.KeepSourceFormatting);

        // -------------------------------------------------
        // 8. Collect paragraphs that lie strictly between the start and end paragraphs.
        // -------------------------------------------------
        bool withinRange = false;
        int extractedCount = 0;
        foreach (Paragraph para in loaded.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para == startParagraph)
            {
                withinRange = true;
                continue; // Skip the start paragraph itself.
            }

            if (para == endParagraph)
                break; // Stop before the end paragraph.

            if (withinRange)
            {
                Node imported = importer.ImportNode(para, true);
                resultBody.AppendChild(imported);
                extractedCount++;
            }
        }

        if (extractedCount == 0)
            throw new InvalidOperationException("No nodes were extracted between the specified markers.");

        // -------------------------------------------------
        // 9. Save the extracted content.
        // -------------------------------------------------
        const string resultPath = "extracted.docx";
        result.Save(resultPath);

        // Simple validation that the output file exists.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("Extraction failed: output file was not created.");
    }
}
