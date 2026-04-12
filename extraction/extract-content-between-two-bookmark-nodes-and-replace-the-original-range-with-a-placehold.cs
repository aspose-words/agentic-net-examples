using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a sample document with two empty bookmarks surrounding some content.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Paragraph before bookmarks.");

        // Insert start bookmark (empty).
        builder.StartBookmark("Start");
        builder.EndBookmark("Start");

        // Content that will be extracted.
        builder.Writeln("First paragraph inside range.");
        builder.Writeln("Second paragraph inside range.");

        // Insert end bookmark (empty).
        builder.StartBookmark("End");
        builder.EndBookmark("End");

        builder.Writeln("Paragraph after bookmarks.");

        // Locate the two bookmarks.
        Bookmark startBookmark = sourceDoc.Range.Bookmarks["Start"];
        Bookmark endBookmark = sourceDoc.Range.Bookmarks["End"];
        if (startBookmark == null || endBookmark == null)
            throw new InvalidOperationException("Required bookmarks were not found.");

        // Determine the nodes that lie between the end of the start bookmark and the start of the end bookmark.
        Node startNode = startBookmark.BookmarkEnd;
        Node endNode = endBookmark.BookmarkStart;

        List<Node> nodesInRange = new List<Node>();
        for (Node cur = startNode.NextSibling; cur != null && cur != endNode; cur = cur.NextSibling)
        {
            nodesInRange.Add(cur);
        }

        if (nodesInRange.Count == 0)
            throw new InvalidOperationException("No content found between the bookmarks.");

        // Create a new document to hold the extracted content.
        Document extractedDoc = new Document();
        extractedDoc.RemoveAllChildren(); // Remove the default empty section.

        // Build a new section with a body.
        Section newSection = new Section(extractedDoc);
        Body newBody = new Body(extractedDoc);
        newSection.AppendChild(newBody);
        extractedDoc.AppendChild(newSection);

        // Import and clone the nodes into the new document, handling inline nodes.
        NodeImporter importer = new NodeImporter(sourceDoc, extractedDoc, ImportFormatMode.KeepSourceFormatting);
        foreach (Node node in nodesInRange)
        {
            Node importedNode = importer.ImportNode(node, true);
            if (importedNode.NodeType == NodeType.Paragraph || importedNode.NodeType == NodeType.Table)
            {
                newBody.AppendChild(importedNode);
            }
            else
            {
                // Wrap inline nodes (e.g., Run) inside a paragraph belonging to the destination document.
                Paragraph wrapper = new Paragraph(extractedDoc);
                wrapper.AppendChild(importedNode);
                newBody.AppendChild(wrapper);
            }
        }

        // Save the extracted content.
        string extractedPath = Path.Combine(Directory.GetCurrentDirectory(), "Extracted.docx");
        extractedDoc.Save(extractedPath);

        // Remove the original nodes from the source document.
        foreach (Node node in nodesInRange)
        {
            node.Remove();
        }

        // Insert a placeholder paragraph after the paragraph that contains the start bookmark.
        Paragraph startParagraph = startBookmark.BookmarkEnd.ParentNode as Paragraph;
        if (startParagraph == null)
            throw new InvalidOperationException("Start bookmark is not inside a paragraph.");

        Paragraph placeholder = new Paragraph(sourceDoc);
        Run placeholderRun = new Run(sourceDoc, "Placeholder content");
        placeholder.AppendChild(placeholderRun);

        // Insert the placeholder as a sibling block.
        CompositeNode body = startParagraph.ParentNode as CompositeNode;
        if (body == null)
            throw new InvalidOperationException("Unable to locate the body to insert the placeholder.");

        body.InsertAfter(placeholder, startParagraph);

        // Save the modified source document.
        string modifiedPath = Path.Combine(Directory.GetCurrentDirectory(), "Modified.docx");
        sourceDoc.Save(modifiedPath);

        // Validate that the output files were created.
        if (!File.Exists(extractedPath))
            throw new FileNotFoundException("Extracted document was not created.", extractedPath);
        if (!File.Exists(modifiedPath))
            throw new FileNotFoundException("Modified document was not created.", modifiedPath);
    }
}
