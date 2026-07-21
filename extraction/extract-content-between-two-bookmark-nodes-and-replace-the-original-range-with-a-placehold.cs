using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // Create a sample document with two bookmarks that enclose some content.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Paragraph before the range.");

        builder.StartBookmark("StartRange");
        builder.Writeln("First paragraph inside range.");
        builder.Writeln("Second paragraph inside range.");
        builder.EndBookmark("StartRange");

        builder.StartBookmark("EndRange");
        builder.Writeln("Third paragraph inside range.");
        builder.EndBookmark("EndRange");

        builder.Writeln("Paragraph after the range.");

        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // Load the document for processing.
        // -----------------------------------------------------------------
        Document doc = new Document(sourcePath);

        // Retrieve the two bookmarks.
        Bookmark startBookmark = doc.Range.Bookmarks["StartRange"];
        Bookmark endBookmark = doc.Range.Bookmarks["EndRange"];
        if (startBookmark == null || endBookmark == null)
            throw new InvalidOperationException("Required bookmarks were not found.");

        // -----------------------------------------------------------------
        // Extract the content between the two bookmarks into a new document.
        // -----------------------------------------------------------------
        // Get the paragraphs that contain the start and end bookmark markers.
        Paragraph startParagraph = startBookmark.BookmarkStart.GetAncestor(NodeType.Paragraph) as Paragraph;
        Paragraph endParagraph = endBookmark.BookmarkEnd.GetAncestor(NodeType.Paragraph) as Paragraph;
        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Unable to locate paragraph boundaries.");

        // Prepare the destination document.
        Document extracted = new Document();
        extracted.RemoveAllChildren();
        Section extractedSection = new Section(extracted);
        extracted.AppendChild(extractedSection);
        Body extractedBody = new Body(extracted);
        extractedSection.AppendChild(extractedBody);

        // Import each paragraph from the source range into the new document.
        NodeImporter importer = new NodeImporter(doc, extracted, ImportFormatMode.KeepSourceFormatting);
        Node current = startParagraph;
        while (current != null)
        {
            Node imported = importer.ImportNode(current, true);
            extractedBody.AppendChild(imported);
            if (current == endParagraph)
                break;
            current = current.NextSibling;
        }

        const string extractedPath = "extracted.docx";
        extracted.Save(extractedPath);
        if (!File.Exists(extractedPath))
            throw new InvalidOperationException("Extraction failed – output file not created.");

        // -----------------------------------------------------------------
        // Replace the original bookmarked range with a placeholder paragraph.
        // -----------------------------------------------------------------
        // Determine the body that holds the range and the paragraph preceding the range.
        Body body = startParagraph.ParentNode as Body;
        Paragraph previousParagraph = startParagraph.PreviousSibling as Paragraph;

        // Remove the paragraphs that belong to the range.
        current = startParagraph;
        while (current != null)
        {
            Node next = current.NextSibling;
            current.Remove();
            if (current == endParagraph)
                break;
            current = next;
        }

        // Insert a placeholder paragraph at the correct position.
        Paragraph placeholder = new Paragraph(doc);
        placeholder.AppendChild(new Run(doc, "[Extracted content removed]"));

        if (previousParagraph != null)
            body.InsertAfter(placeholder, previousParagraph);
        else
            body.PrependChild(placeholder);

        // Save the modified document.
        const string modifiedPath = "modified.docx";
        doc.Save(modifiedPath);
        if (!File.Exists(modifiedPath))
            throw new InvalidOperationException("Modified document was not saved.");
    }
}
