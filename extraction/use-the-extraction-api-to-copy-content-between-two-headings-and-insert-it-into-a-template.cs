using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a source document with two headings and some content between them.
        const string sourcePath = "source.docx";
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

        // First heading.
        srcBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        srcBuilder.Writeln("Start");

        // Content to be extracted.
        srcBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        srcBuilder.Writeln("Paragraph 1 between headings.");
        srcBuilder.Writeln("Paragraph 2 between headings.");

        // Insert a simple table.
        srcBuilder.StartTable();
        srcBuilder.InsertCell();
        srcBuilder.Write("Cell A1");
        srcBuilder.InsertCell();
        srcBuilder.Write("Cell B1");
        srcBuilder.EndRow();
        srcBuilder.EndTable();

        // Second heading.
        srcBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        srcBuilder.Writeln("End");

        // Save the source document.
        sourceDoc.Save(sourcePath);

        // Load the source document.
        Document loadedSource = new Document(sourcePath);

        // Locate the start and end heading paragraphs.
        Paragraph startHeading = FindHeadingParagraph(loadedSource, "Start");
        Paragraph endHeading = FindHeadingParagraph(loadedSource, "End");

        if (startHeading == null || endHeading == null)
            throw new InvalidOperationException("Required headings were not found in the source document.");

        // Collect all nodes that are between the two headings (exclusive).
        List<Node> nodesBetween = new List<Node>();
        Node current = startHeading.NextSibling;
        while (current != null && current != endHeading)
        {
            nodesBetween.Add(current);
            current = current.NextSibling;
        }

        if (nodesBetween.Count == 0)
            throw new InvalidOperationException("No content found between the specified headings.");

        // Create a template document with a bookmark where the extracted content will be inserted.
        const string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder tmplBuilder = new DocumentBuilder(templateDoc);
        tmplBuilder.StartBookmark("InsertHere");
        tmplBuilder.Writeln("[Placeholder]");
        tmplBuilder.EndBookmark("InsertHere");
        templateDoc.Save(templatePath);

        // Load the template document.
        Document loadedTemplate = new Document(templatePath);
        Bookmark insertBookmark = loadedTemplate.Range.Bookmarks["InsertHere"];
        if (insertBookmark == null)
            throw new InvalidOperationException("Insert bookmark not found in the template document.");

        // The bookmark is inside a paragraph; we'll insert after that paragraph.
        Paragraph bookmarkParagraph = insertBookmark.BookmarkStart.ParentNode as Paragraph;
        if (bookmarkParagraph == null)
            throw new InvalidOperationException("Bookmark is not placed inside a paragraph.");

        // Prepare a NodeImporter for efficient importing.
        NodeImporter importer = new NodeImporter(loadedSource, loadedTemplate, ImportFormatMode.KeepSourceFormatting);

        // Insert the extracted nodes into the template.
        Node previousNode = bookmarkParagraph;
        foreach (Node node in nodesBetween)
        {
            Node importedNode = importer.ImportNode(node, true);
            bookmarkParagraph.ParentNode.InsertAfter(importedNode, previousNode);
            previousNode = importedNode;
        }

        // Optionally remove the placeholder paragraph that contained the bookmark.
        bookmarkParagraph.Remove();

        // Save the resulting document.
        const string resultPath = "result.docx";
        loadedTemplate.Save(resultPath);

        // Verify that the result file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("Result document was not created.");

        // Clean up temporary files (optional).
        // File.Delete(sourcePath);
        // File.Delete(templatePath);
    }

    // Helper method to find a heading paragraph with the specified text.
    private static Paragraph FindHeadingParagraph(Document doc, string headingText)
    {
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // Ensure the paragraph uses a heading style.
            if (para.ParagraphFormat.IsHeading)
            {
                // GetText includes the paragraph break; trim it for comparison.
                if (para.GetText().Trim() == headingText)
                    return para;
            }
        }
        return null;
    }
}
