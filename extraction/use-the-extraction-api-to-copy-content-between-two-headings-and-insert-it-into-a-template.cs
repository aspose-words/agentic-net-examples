using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class ExtractionBetweenHeadings
{
    public static void Main()
    {
        // ---------- Create source document ----------
        Document source = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(source);
        srcBuilder.Writeln("Document Title");

        srcBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        srcBuilder.Writeln("Heading 1");

        srcBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        srcBuilder.Writeln("Paragraph A under Heading 1.");
        srcBuilder.Writeln("Paragraph B under Heading 1.");

        srcBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        srcBuilder.Writeln("Heading 2");

        srcBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        srcBuilder.Writeln("Paragraph C under Heading 2.");

        source.Save("source.docx");

        // ---------- Create template document with a bookmark ----------
        Document template = new Document();
        DocumentBuilder tmplBuilder = new DocumentBuilder(template);
        tmplBuilder.Writeln("Template Header");
        tmplBuilder.StartBookmark("InsertHere");
        tmplBuilder.Writeln("[Placeholder]");
        tmplBuilder.EndBookmark("InsertHere");
        template.Save("template.docx");

        // ---------- Load documents ----------
        Document srcDoc = new Document("source.docx");
        Document tmplDoc = new Document("template.docx");

        // ---------- Locate start and end headings ----------
        Paragraph startHeading = null;
        Paragraph endHeading = null;

        foreach (Paragraph para in srcDoc.FirstSection.Body.Paragraphs)
        {
            if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
            {
                string text = para.GetText().Trim();
                if (text == "Heading 1")
                    startHeading = para;
                else if (text == "Heading 2")
                    endHeading = para;
            }
        }

        if (startHeading == null || endHeading == null)
            throw new InvalidOperationException("Required headings were not found in the source document.");

        // ---------- Collect nodes between the two headings ----------
        List<Node> nodesToCopy = new List<Node>();
        Node current = startHeading.NextSibling;
        while (current != null && current != endHeading)
        {
            nodesToCopy.Add(current);
            current = current.NextSibling;
        }

        if (nodesToCopy.Count == 0)
            throw new InvalidOperationException("No content found between the specified headings.");

        // ---------- Locate the bookmark in the template ----------
        Bookmark insertBookmark = tmplDoc.Range.Bookmarks["InsertHere"];
        if (insertBookmark == null)
            throw new InvalidOperationException("Insert bookmark not found in the template document.");

        Paragraph bookmarkParagraph = insertBookmark.BookmarkStart.ParentNode as Paragraph;
        if (bookmarkParagraph == null)
            throw new InvalidOperationException("Bookmark is not placed inside a paragraph.");

        // ---------- Remove the placeholder run inside the bookmark paragraph ----------
        foreach (Run run in bookmarkParagraph.Runs)
        {
            if (run.Text.Contains("[Placeholder]"))
            {
                run.Remove();
                break;
            }
        }

        // ---------- Import and insert cloned nodes ----------
        NodeImporter importer = new NodeImporter(srcDoc, tmplDoc, ImportFormatMode.KeepSourceFormatting);
        Node insertionPoint = bookmarkParagraph;

        foreach (Node node in nodesToCopy)
        {
            Node importedNode = importer.ImportNode(node, true);
            insertionPoint.ParentNode.InsertAfter(importedNode, insertionPoint);
            insertionPoint = importedNode;
        }

        // ---------- Save the resulting document ----------
        tmplDoc.Save("result.docx");

        // ---------- Validate output ----------
        if (!File.Exists("result.docx"))
            throw new InvalidOperationException("Result document was not created.");
    }
}
