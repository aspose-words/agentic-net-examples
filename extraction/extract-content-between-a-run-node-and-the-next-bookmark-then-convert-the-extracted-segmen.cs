using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // -------------------------------------------------
        // 1. Build a sample document containing a run and a bookmark.
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Paragraph before the run.");

        // Paragraph that will contain the target Run.
        builder.Writeln("Paragraph with the target run:");
        Paragraph paraWithRun = builder.CurrentParagraph;
        Run targetRun = new Run(sourceDoc, "TargetRun");
        paraWithRun.AppendChild(targetRun);

        // Some additional content after the run.
        builder.Writeln("Text after the run, before the bookmark.");

        // Insert a bookmark after the above content.
        builder.StartBookmark("MyBookmark");
        builder.Writeln("Content inside the bookmark.");
        builder.EndBookmark("MyBookmark");

        // Save the source document (optional, for inspection).
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // -------------------------------------------------
        // 2. Locate the first bookmark that appears after the target run.
        // -------------------------------------------------
        BookmarkStart nextBookmarkStart = null;
        Node node = targetRun;
        while ((node = node.NextPreOrder(sourceDoc)) != null)
        {
            if (node.NodeType == NodeType.BookmarkStart)
            {
                nextBookmarkStart = (BookmarkStart)node;
                break;
            }
        }

        if (nextBookmarkStart == null)
            throw new InvalidOperationException("No bookmark found after the target run.");

        // -------------------------------------------------
        // 3. Extract all nodes that lie between the run and the bookmark.
        //    Only block‑level nodes (Paragraph, Table) can be added directly to a Body.
        //    Inline nodes (Run) are wrapped in a Paragraph first.
        // -------------------------------------------------
        Document extractedDoc = new Document(); // contains a default section & body.
        NodeImporter importer = new NodeImporter(sourceDoc, extractedDoc, ImportFormatMode.KeepSourceFormatting);
        Node current = targetRun;

        while ((current = current.NextPreOrder(sourceDoc)) != null && current != nextBookmarkStart)
        {
            // Import the node into the destination document.
            Node imported = importer.ImportNode(current, true);

            switch (imported.NodeType)
            {
                case NodeType.Paragraph:
                case NodeType.Table:
                    // Block nodes can be appended directly.
                    extractedDoc.FirstSection.Body.AppendChild(imported);
                    break;

                case NodeType.Run:
                    // Inline nodes must be placed inside a paragraph.
                    Paragraph wrapper = new Paragraph(extractedDoc);
                    wrapper.AppendChild(imported);
                    extractedDoc.FirstSection.Body.AppendChild(wrapper);
                    break;

                // Skip other node types (e.g., BookmarkStart/End) that are not needed for the output.
                default:
                    break;
            }
        }

        // -------------------------------------------------
        // 4. Save the extracted fragment as HTML.
        // -------------------------------------------------
        const string htmlPath = "extracted.html";
        extractedDoc.Save(htmlPath, SaveFormat.Html);

        // Verify that the HTML file was created.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML extraction output was not created.");
    }
}
