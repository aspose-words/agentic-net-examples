using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // ------------------------------------------------------------
        // 1. Create a sample source document.
        // ------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Paragraph before.");
        builder.Write("Run before bookmark "); // This creates a Run node.
        builder.Writeln(); // End the paragraph.

        // Content that should be extracted (between the run and the bookmark).
        builder.Writeln("Paragraph between run and bookmark.");

        // Bookmark that follows the run.
        builder.StartBookmark("MyBookmark");
        builder.Writeln("Inside bookmark.");
        builder.EndBookmark("MyBookmark");

        builder.Writeln("Paragraph after bookmark.");

        const string sourcePath = "sample.docx";
        sourceDoc.Save(sourcePath);

        // ------------------------------------------------------------
        // 2. Load the document for processing.
        // ------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // Locate the specific Run node that precedes the bookmark.
        Run targetRun = loadedDoc.GetChildNodes(NodeType.Run, true)
                                 .OfType<Run>()
                                 .FirstOrDefault(r => r.GetText().Contains("Run before bookmark"));
        if (targetRun == null)
            throw new InvalidOperationException("Target run not found.");

        // Find the next BookmarkStart node after the run in document order.
        NodeCollection allNodes = loadedDoc.GetChildNodes(NodeType.Any, true);
        int runIndex = -1;
        for (int i = 0; i < allNodes.Count; i++)
        {
            if (allNodes[i] == targetRun)
            {
                runIndex = i;
                break;
            }
        }
        if (runIndex == -1)
            throw new InvalidOperationException("Run position could not be determined.");

        BookmarkStart nextBookmarkStart = null;
        for (int i = runIndex + 1; i < allNodes.Count; i++)
        {
            if (allNodes[i].NodeType == NodeType.BookmarkStart)
            {
                nextBookmarkStart = (BookmarkStart)allNodes[i];
                break;
            }
        }
        if (nextBookmarkStart == null)
            throw new InvalidOperationException("No bookmark found after the target run.");

        // ------------------------------------------------------------
        // 3. Determine the paragraphs that bound the extraction range.
        // ------------------------------------------------------------
        Paragraph startParagraph = targetRun.ParentNode as Paragraph;
        Paragraph endParagraph = nextBookmarkStart.ParentNode as Paragraph;
        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Unable to locate bounding paragraphs.");

        Body sourceBody = startParagraph.ParentNode as Body;
        if (sourceBody == null)
            throw new InvalidOperationException("Source body not found.");

        // ------------------------------------------------------------
        // 4. Create a new document and import the paragraphs that lie between the bounds.
        // ------------------------------------------------------------
        Document extractedDoc = new Document();
        extractedDoc.RemoveAllChildren();

        Section resultSection = new Section(extractedDoc);
        extractedDoc.AppendChild(resultSection);
        Body resultBody = new Body(extractedDoc);
        resultSection.AppendChild(resultBody);

        // Importer to copy nodes from sourceDoc to extractedDoc.
        NodeImporter importer = new NodeImporter(loadedDoc, extractedDoc, ImportFormatMode.KeepSourceFormatting);

        // Walk through the sibling paragraphs after startParagraph up to (but not including) endParagraph.
        Node currentNode = startParagraph.NextSibling;
        while (currentNode != null && currentNode != endParagraph)
        {
            if (currentNode.NodeType == NodeType.Paragraph)
            {
                Paragraph para = (Paragraph)currentNode;
                Paragraph importedPara = (Paragraph)importer.ImportNode(para, true);
                resultBody.AppendChild(importedPara);
            }
            else if (currentNode.NodeType == NodeType.Table)
            {
                // If a table appears in the range, import it as well.
                Aspose.Words.Tables.Table table = (Aspose.Words.Tables.Table)currentNode;
                Aspose.Words.Tables.Table importedTable = (Aspose.Words.Tables.Table)importer.ImportNode(table, true);
                resultBody.AppendChild(importedTable);
            }
            // Advance to the next sibling.
            currentNode = currentNode.NextSibling;
        }

        // ------------------------------------------------------------
        // 5. Save the extracted segment as HTML.
        // ------------------------------------------------------------
        const string htmlPath = "extracted.html";
        extractedDoc.Save(htmlPath, SaveFormat.Html);

        // Verify that the HTML file was created.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML extraction output was not created.");
    }
}
