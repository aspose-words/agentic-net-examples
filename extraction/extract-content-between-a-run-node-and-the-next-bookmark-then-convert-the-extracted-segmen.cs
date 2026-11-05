using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // ------------------------------------------------------------
        // 1. Create a sample source document with a run and a bookmark.
        // ------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Paragraph before runs.");
        builder.Write("RunA ");                     // First run (not target)
        builder.Write("RunB ");                     // Target run
        builder.Write(" TextBetween ");             // Content we want to extract
        builder.StartBookmark("TargetBookmark");    // Bookmark after the content
        builder.Write("Bookmarked content");
        builder.EndBookmark("TargetBookmark");
        builder.Writeln(); // End of paragraph

        // Save the source document (optional, for inspection).
        sourceDoc.Save("source.docx");

        // ------------------------------------------------------------
        // 2. Load the document as if it were an external file.
        // ------------------------------------------------------------
        Document doc = new Document("source.docx");

        // ------------------------------------------------------------
        // 3. Locate the target Run node that contains "RunB ".
        // ------------------------------------------------------------
        Run targetRun = null;
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            if (run.Text != null && run.Text.Contains("RunB"))
            {
                targetRun = run;
                break;
            }
        }

        if (targetRun == null)
            throw new InvalidOperationException("Target run not found.");

        // ------------------------------------------------------------
        // 4. Locate the bookmark that follows the target run.
        // ------------------------------------------------------------
        BookmarkStart nextBookmarkStart = null;
        foreach (BookmarkStart bmStart in doc.GetChildNodes(NodeType.BookmarkStart, true))
        {
            if (bmStart.Name == "TargetBookmark")
            {
                nextBookmarkStart = bmStart;
                break;
            }
        }

        if (nextBookmarkStart == null)
            throw new InvalidOperationException("Target bookmark not found.");

        // ------------------------------------------------------------
        // 5. Prepare a new document that will hold the extracted content.
        // ------------------------------------------------------------
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren();

        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);

        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // The extracted inline runs will be placed inside this paragraph.
        Paragraph extractedParagraph = new Paragraph(resultDoc);
        resultBody.AppendChild(extractedParagraph);

        // ------------------------------------------------------------
        // 6. Extract nodes that appear after the target run up to (but not including) the bookmark.
        // ------------------------------------------------------------
        NodeImporter importer = new NodeImporter(doc, resultDoc, ImportFormatMode.KeepSourceFormatting);

        Node currentNode = targetRun.NextSibling;
        while (currentNode != null && !(currentNode is BookmarkStart bs && bs.Name == "TargetBookmark"))
        {
            // Preserve reference to the next sibling before we import the current node.
            Node nextNode = currentNode.NextSibling;

            // Import the node into the destination document.
            Node importedNode = importer.ImportNode(currentNode, true);

            // Decide where to place the imported node.
            if (importedNode is Run importedRun)
            {
                // Inline run – add to the paragraph we created.
                extractedParagraph.AppendChild(importedRun);
            }
            else if (importedNode is Paragraph importedParagraph)
            {
                // A full paragraph – add directly to the body.
                resultBody.AppendChild(importedParagraph);
            }
            else if (importedNode is Table importedTable)
            {
                // Table – also a block node, add to the body.
                resultBody.AppendChild(importedTable);
            }
            else
            {
                // For any other node types, attempt to add to the body.
                resultBody.AppendChild(importedNode);
            }

            currentNode = nextNode;
        }

        // ------------------------------------------------------------
        // 7. Validate that something was extracted.
        // ------------------------------------------------------------
        bool hasRuns = extractedParagraph.Runs.Count > 0;
        bool hasOtherBlocks = resultBody.Paragraphs.Count > 1 || resultBody.Tables.Count > 0;
        if (!hasRuns && !hasOtherBlocks)
            throw new InvalidOperationException("No content was extracted between the run and the bookmark.");

        // ------------------------------------------------------------
        // 8. Save the extracted segment as HTML.
        // ------------------------------------------------------------
        string htmlPath = "extracted.html";
        resultDoc.Save(htmlPath, SaveFormat.Html);

        // ------------------------------------------------------------
        // 9. Verify that the HTML file was created.
        // ------------------------------------------------------------
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML output was not created.");

        // Cleanup (optional – comment out if you want to inspect the files).
        // File.Delete("source.docx");
        // File.Delete(htmlPath);
    }
}
