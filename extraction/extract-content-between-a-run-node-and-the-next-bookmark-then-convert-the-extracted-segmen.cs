using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // 1. Create a sample source document with runs and a bookmark.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Paragraph before the target run.");
        builder.Write("FirstRun");               // This creates a Run inside a Paragraph.
        builder.Writeln();                       // End the paragraph.

        builder.Writeln("Paragraph between run and bookmark.");
        builder.StartBookmark("TargetBookmark");
        builder.Writeln("Content inside the bookmark.");
        builder.EndBookmark("TargetBookmark");
        builder.Writeln("Paragraph after the bookmark.");

        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // 2. Locate the first Run node in the document.
        Run startRun = sourceDoc.GetChildNodes(NodeType.Run, true)[0] as Run;
        if (startRun == null)
            throw new InvalidOperationException("Run node not found.");

        // 3. Find the next BookmarkStart node after the Run.
        BookmarkStart nextBookmark = null;
        Node traversalNode = startRun;
        while ((traversalNode = traversalNode.NextPreOrder(sourceDoc)) != null)
        {
            if (traversalNode.NodeType == NodeType.BookmarkStart)
            {
                nextBookmark = (BookmarkStart)traversalNode;
                break;
            }
        }

        if (nextBookmark == null)
            throw new InvalidOperationException("Next bookmark not found.");

        // 4. Extract content between the Run and the next BookmarkStart.
        Document extractedDoc = new Document();
        extractedDoc.RemoveAllChildren(); // Ensure a clean document.

        // Build a minimal document structure: Section -> Body.
        Section section = new Section(extractedDoc);
        extractedDoc.AppendChild(section);
        Body body = new Body(extractedDoc);
        section.AppendChild(body);

        // Use NodeImporter to import nodes from sourceDoc into extractedDoc.
        NodeImporter importer = new NodeImporter(sourceDoc, extractedDoc, ImportFormatMode.KeepSourceFormatting);

        // Start traversal from the node immediately after the startRun.
        Node currentNode = startRun;
        while ((currentNode = currentNode.NextPreOrder(sourceDoc)) != null)
        {
            if (currentNode == nextBookmark)
                break; // Stop before the bookmark.

            // Import only block-level nodes that can be children of Body.
            if (currentNode.NodeType == NodeType.Paragraph || currentNode.NodeType == NodeType.Table)
            {
                Node importedNode = importer.ImportNode(currentNode, true);
                body.AppendChild(importedNode);
            }
        }

        // 5. Convert the extracted segment to HTML.
        const string htmlPath = "extracted.html";
        extractedDoc.Save(htmlPath, SaveFormat.Html);

        // Validate that the HTML file was created.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML extraction output was not created.");
    }
}
