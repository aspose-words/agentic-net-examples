using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class BookmarkRangeExtraction
{
    public static void Main()
    {
        // Create a sample document with two bookmarks surrounding some content.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Paragraph before bookmarks.");

        // Start of the range to extract.
        builder.StartBookmark("StartRange");
        builder.Writeln("First line inside range.");
        builder.Writeln("Second line inside range.");

        // Add a table inside the range.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        builder.EndBookmark("StartRange");

        // End of the range to extract.
        builder.StartBookmark("EndRange");
        builder.Writeln("Third line inside range.");
        builder.EndBookmark("EndRange");

        builder.Writeln("Paragraph after bookmarks.");

        // Save the source document.
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // Load the document (could also use the same instance).
        Document doc = new Document(sourcePath);

        // Retrieve the two bookmarks.
        Bookmark startBookmark = doc.Range.Bookmarks["StartRange"];
        Bookmark endBookmark = doc.Range.Bookmarks["EndRange"];
        if (startBookmark == null || endBookmark == null)
            throw new InvalidOperationException("Required bookmarks were not found.");

        // Determine the block-level nodes that lie between the two bookmark boundaries.
        Paragraph startParagraph = startBookmark.BookmarkStart.ParentNode as Paragraph;
        Paragraph endParagraph = endBookmark.BookmarkEnd.ParentNode as Paragraph;
        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Bookmarks are not placed inside paragraphs.");

        List<Node> nodesInRange = new List<Node>();
        for (Node cur = startParagraph; cur != null; cur = cur.NextSibling)
        {
            nodesInRange.Add(cur);
            if (cur == endParagraph)
                break;
        }

        if (nodesInRange.Count == 0)
            throw new InvalidOperationException("No content found between the specified bookmarks.");

        // Create a new document to hold the extracted content.
        Document extractedDoc = new Document();
        extractedDoc.RemoveAllChildren();
        Section extractedSection = new Section(extractedDoc);
        extractedDoc.AppendChild(extractedSection);
        Body extractedBody = new Body(extractedDoc);
        extractedSection.AppendChild(extractedBody);

        // Import the nodes into the extracted document.
        NodeImporter importer = new NodeImporter(doc, extractedDoc, ImportFormatMode.KeepSourceFormatting);
        foreach (Node node in nodesInRange)
        {
            Node importedNode = importer.ImportNode(node, true);
            extractedBody.AppendChild(importedNode);
        }

        // Save the extracted content.
        const string extractedPath = "extracted.docx";
        extractedDoc.Save(extractedPath);
        if (!File.Exists(extractedPath))
            throw new InvalidOperationException("Extracted document was not created.");

        // Remove the original nodes from the source document.
        // Capture the parent body and the node that will precede the placeholder.
        Body parentBody = startParagraph.ParentNode as Body;
        Node referenceNode = startParagraph.PreviousSibling; // may be null

        foreach (Node node in nodesInRange)
        {
            node.Remove();
        }

        // Insert a placeholder paragraph at the original location.
        Paragraph placeholder = new Paragraph(doc);
        placeholder.AppendChild(new Run(doc, "[Extracted content removed]"));
        if (referenceNode == null)
        {
            // Insert at the beginning of the body.
            parentBody.PrependChild(placeholder);
        }
        else
        {
            parentBody.InsertAfter(placeholder, referenceNode);
        }

        // Save the modified original document.
        const string modifiedPath = "modified.docx";
        doc.Save(modifiedPath);
        if (!File.Exists(modifiedPath))
            throw new InvalidOperationException("Modified document was not created.");

        // Simple validation: ensure the placeholder exists in the modified document.
        Document validationDoc = new Document(modifiedPath);
        bool placeholderFound = false;
        foreach (Paragraph para in validationDoc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.GetText().Contains("[Extracted content removed]"))
            {
                placeholderFound = true;
                break;
            }
        }

        if (!placeholderFound)
            throw new InvalidOperationException("Placeholder paragraph was not inserted correctly.");

        Console.WriteLine("Extraction and replacement completed.");
    }
}
