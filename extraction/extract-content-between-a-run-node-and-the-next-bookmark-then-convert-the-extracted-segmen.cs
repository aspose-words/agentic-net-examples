using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a run and a following bookmark.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Paragraph before the run.");
        builder.Write("RunText"); // This creates a Run inside the current paragraph.
        builder.Writeln(); // End the paragraph.

        // Insert a bookmark after the run.
        builder.StartBookmark("TargetBookmark");
        builder.Writeln("Text inside the bookmark.");
        builder.EndBookmark("TargetBookmark");

        // Save the source document (optional, just to have a file on disk).
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // Load the document back.
        Document loadedDoc = new Document(sourcePath);

        // Locate the first Run node.
        Run runNode = loadedDoc.GetChildNodes(NodeType.Run, true)[0] as Run;
        if (runNode == null)
            throw new InvalidOperationException("Run node not found.");

        // Find the next BookmarkStart node after the run using pre‑order traversal.
        BookmarkStart nextBookmark = null;
        Node traversal = runNode;
        while ((traversal = traversal.NextPreOrder(null)) != null)
        {
            if (traversal.NodeType == NodeType.BookmarkStart)
            {
                nextBookmark = (BookmarkStart)traversal;
                break;
            }
        }

        if (nextBookmark == null)
            throw new InvalidOperationException("Next bookmark not found.");

        // Build a new document that will contain the extracted content.
        Document extractedDoc = new Document();
        extractedDoc.RemoveAllChildren();

        Section section = new Section(extractedDoc);
        extractedDoc.AppendChild(section);

        Body body = new Body(extractedDoc);
        section.AppendChild(body);

        // Collect nodes that lie between the run and the bookmark (exclusive).
        Node current = runNode.NextPreOrder(null);
        bool anyNodeAdded = false;

        while (current != null && current != nextBookmark)
        {
            // Import the node into the destination document.
            Node imported = extractedDoc.ImportNode(current, true);

            // Append only nodes that are valid children of Body.
            if (imported.NodeType == NodeType.Paragraph || imported.NodeType == NodeType.Table)
            {
                body.AppendChild(imported);
                anyNodeAdded = true;
            }
            else if (imported.NodeType == NodeType.Run)
            {
                // Wrap isolated runs in a paragraph.
                Paragraph para = new Paragraph(extractedDoc);
                para.AppendChild(imported);
                body.AppendChild(para);
                anyNodeAdded = true;
            }

            current = current.NextPreOrder(null);
        }

        if (!anyNodeAdded)
            throw new InvalidOperationException("No content was extracted between the run and the bookmark.");

        // Convert the extracted segment to HTML.
        string htmlContent = extractedDoc.ToString(SaveFormat.Html);

        // Save the HTML to a file for verification.
        const string htmlPath = "extracted.html";
        File.WriteAllText(htmlPath, htmlContent);

        Console.WriteLine($"Extraction completed. HTML saved to '{htmlPath}'.");
    }
}
