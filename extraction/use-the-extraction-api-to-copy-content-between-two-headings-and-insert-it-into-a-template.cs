using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a source document with two headings and some content between them.
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

        // First heading (start marker).
        srcBuilder.ParagraphFormat.StyleName = "Heading 1";
        srcBuilder.Writeln("Start Heading");

        // Content that will be extracted.
        srcBuilder.ParagraphFormat.StyleName = "Normal";
        srcBuilder.Writeln("This is the first paragraph between headings.");
        srcBuilder.Writeln("This is the second paragraph between headings.");

        // Second heading (end marker).
        srcBuilder.ParagraphFormat.StyleName = "Heading 1";
        srcBuilder.Writeln("End Heading");

        // Additional content after the end heading (should not be extracted).
        srcBuilder.ParagraphFormat.StyleName = "Normal";
        srcBuilder.Writeln("Content after the end heading.");

        // Locate the start and end heading paragraphs.
        Paragraph startHeading = FindParagraphByText(sourceDoc, "Start Heading");
        Paragraph endHeading = FindParagraphByText(sourceDoc, "End Heading");

        if (startHeading == null || endHeading == null)
            throw new InvalidOperationException("Required headings were not found in the source document.");

        // Collect all nodes that lie between the two headings (exclusive).
        List<Node> nodesBetween = new List<Node>();
        Node curNode = startHeading.NextSibling;
        while (curNode != null && curNode != endHeading)
        {
            nodesBetween.Add(curNode);
            curNode = curNode.NextSibling;
        }

        if (nodesBetween.Count == 0)
            throw new InvalidOperationException("No content found between the specified headings.");

        // Create a template document with a bookmark where the extracted content will be inserted.
        Document templateDoc = new Document();
        DocumentBuilder tmplBuilder = new DocumentBuilder(templateDoc);
        tmplBuilder.Writeln("Template Document Header");
        tmplBuilder.StartBookmark("InsertHere");
        tmplBuilder.Writeln("[Placeholder for extracted content]");
        tmplBuilder.EndBookmark("InsertHere");
        tmplBuilder.Writeln("Template Document Footer");

        // Prepare a NodeImporter for efficient node copying.
        NodeImporter importer = new NodeImporter(sourceDoc, templateDoc, ImportFormatMode.KeepSourceFormatting);

        // Find the insertion point (the bookmark start's parent paragraph).
        Bookmark insertBookmark = templateDoc.Range.Bookmarks["InsertHere"];
        Node insertionPoint = insertBookmark.BookmarkStart.ParentNode; // This is a Paragraph.
        CompositeNode parent = insertionPoint.ParentNode; // The Body that contains the paragraph.

        // Insert each extracted node after the insertion point, preserving order.
        foreach (Node node in nodesBetween)
        {
            Node importedNode = importer.ImportNode(node, true);
            parent.InsertAfter(importedNode, insertionPoint);
            insertionPoint = importedNode;
        }

        // Optionally, remove the placeholder paragraph that contained the bookmark text.
        // The bookmark itself is empty after insertion, so we can delete its parent paragraph.
        Paragraph placeholderParagraph = (Paragraph)insertBookmark.BookmarkStart.ParentNode;
        placeholderParagraph.Remove();

        // Save the resulting document.
        string outputPath = "Result.docx";
        templateDoc.Save(outputPath);
        Console.WriteLine($"Extracted content inserted and saved to '{outputPath}'.");
    }

    // Helper method to locate a paragraph whose visible text matches the supplied string.
    private static Paragraph FindParagraphByText(Document doc, string text)
    {
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            // GetText includes the paragraph break; Trim it for comparison.
            if (para.GetText().Trim() == text)
                return para;
        }
        return null;
    }
}
