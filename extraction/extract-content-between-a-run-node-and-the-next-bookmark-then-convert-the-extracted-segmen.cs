using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class ExtractionExample
{
    public static void Main()
    {
        // -----------------------------
        // Create a sample source document
        // -----------------------------
        Document sourceDoc = new Document();
        sourceDoc.RemoveAllChildren();

        Section section = new Section(sourceDoc);
        sourceDoc.AppendChild(section);
        Body body = new Body(sourceDoc);
        section.AppendChild(body);

        // Paragraph 1
        Paragraph para1 = new Paragraph(sourceDoc);
        para1.AppendChild(new Run(sourceDoc, "This is the first paragraph. "));
        body.AppendChild(para1);

        // Paragraph 2 – contains the start Run marker
        Paragraph para2 = new Paragraph(sourceDoc);
        Run startRun = new Run(sourceDoc, "StartRun");
        para2.AppendChild(startRun);
        para2.AppendChild(new Run(sourceDoc, " Some text after the start run. "));
        body.AppendChild(para2);

        // Paragraph 3 – content that should be extracted
        Paragraph para3 = new Paragraph(sourceDoc);
        para3.AppendChild(new Run(sourceDoc, "Content that should be extracted. "));
        body.AppendChild(para3);

        // Paragraph 4 – contains the end Bookmark marker
        Paragraph para4 = new Paragraph(sourceDoc);
        BookmarkStart bookmarkStart = new BookmarkStart(sourceDoc, "MyBookmark");
        para4.AppendChild(bookmarkStart);
        para4.AppendChild(new Run(sourceDoc, "Text inside the bookmark."));
        BookmarkEnd bookmarkEnd = new BookmarkEnd(sourceDoc, "MyBookmark");
        para4.AppendChild(bookmarkEnd);
        body.AppendChild(para4);

        // Paragraph 5 – after the bookmark
        Paragraph para5 = new Paragraph(sourceDoc);
        para5.AppendChild(new Run(sourceDoc, "Text after the bookmark."));
        body.AppendChild(para5);

        // -------------------------------------------------
        // Locate the start Run and the next BookmarkStart
        // -------------------------------------------------
        Run firstRun = null;
        NodeCollection runs = sourceDoc.GetChildNodes(NodeType.Run, true);
        if (runs.Count > 0)
        {
            firstRun = (Run)runs[0];
        }

        if (firstRun == null)
        {
            throw new InvalidOperationException("No Run node found in the document.");
        }

        // Find the first BookmarkStart that appears after the start Run in document order
        BookmarkStart nextBookmark = null;
        Node cur = firstRun.NextPreOrder(sourceDoc);
        while (cur != null)
        {
            if (cur.NodeType == NodeType.BookmarkStart)
            {
                nextBookmark = (BookmarkStart)cur;
                break;
            }

            cur = cur.NextPreOrder(sourceDoc);
        }

        if (nextBookmark == null)
        {
            throw new InvalidOperationException("No BookmarkStart found after the Run node.");
        }

        // -------------------------------------------------
        // Extract content between the Run and the Bookmark
        // -------------------------------------------------
        Document extractedDoc = new Document();
        extractedDoc.RemoveAllChildren();
        Section extractedSection = new Section(extractedDoc);
        extractedDoc.AppendChild(extractedSection);
        Body extractedBody = new Body(extractedDoc);
        extractedSection.AppendChild(extractedBody);

        bool anyNodeAdded = false;

        // 1. Handle remaining runs in the same paragraph as the start Run
        Paragraph startParagraph = (Paragraph)firstRun.ParentNode;
        Node sibling = firstRun.NextSibling;
        Paragraph targetParagraph = null;

        while (sibling != null && sibling != nextBookmark)
        {
            if (sibling.NodeType == NodeType.Run)
            {
                if (targetParagraph == null)
                {
                    targetParagraph = new Paragraph(extractedDoc);
                    extractedBody.AppendChild(targetParagraph);
                }

                Node importedRun = extractedDoc.ImportNode(sibling, true, ImportFormatMode.KeepSourceFormatting);
                targetParagraph.AppendChild(importedRun);
                anyNodeAdded = true;
            }

            sibling = sibling.NextSibling;
        }

        // 2. Process whole paragraphs that appear after the start paragraph
        Node paragraphNode = startParagraph.NextSibling;
        while (paragraphNode != null && paragraphNode != nextBookmark.ParentNode)
        {
            if (paragraphNode.NodeType == NodeType.Paragraph)
            {
                Node importedParagraph = extractedDoc.ImportNode(paragraphNode, true, ImportFormatMode.KeepSourceFormatting);
                extractedBody.AppendChild(importedParagraph);
                anyNodeAdded = true;
            }

            paragraphNode = paragraphNode.NextSibling;
        }

        if (!anyNodeAdded)
        {
            throw new InvalidOperationException("No content was found between the Run and the Bookmark.");
        }

        // -------------------------------------------------
        // Save the extracted content as HTML
        // -------------------------------------------------
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Extracted.html");
        extractedDoc.Save(outputPath, SaveFormat.Html);

        Console.WriteLine($"Extracted content saved to: {outputPath}");
    }
}
