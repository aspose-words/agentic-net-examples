using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class ImportCommentsExample
{
    public static void Main()
    {
        // Prepare a sample XML file that contains exported comment data.
        // Each comment includes author, initials, date, text, and the index of the paragraph it belongs to.
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Comments>
  <Comment>
    <Author>John Doe</Author>
    <Initial>JD</Initial>
    <Date>2023-01-01T10:00:00</Date>
    <Text>This is a comment for the first paragraph.</Text>
    <ParagraphIndex>0</ParagraphIndex>
  </Comment>
  <Comment>
    <Author>Jane Smith</Author>
    <Initial>JS</Initial>
    <Date>2023-01-02T11:30:00</Date>
    <Text>Second paragraph needs review.</Text>
    <ParagraphIndex>1</ParagraphIndex>
  </Comment>
</Comments>";
        const string xmlFileName = "comments.xml";
        File.WriteAllText(xmlFileName, xmlContent);

        // Load the XML document that holds the comment information.
        XDocument xDoc = XDocument.Load(xmlFileName);

        // Create a new Word document that will receive the imported comments.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a few paragraphs to the document – these will be the targets for the comments.
        builder.Writeln("First paragraph of the new document.");
        builder.Writeln("Second paragraph of the new document.");
        builder.Writeln("Third paragraph (no comment will be attached).");

        // Process each comment entry from the XML and attach it to the appropriate paragraph.
        foreach (XElement commentElement in xDoc.Root?.Elements("Comment") ?? Enumerable.Empty<XElement>())
        {
            // Extract comment metadata safely.
            string? author = commentElement.Element("Author")?.Value;
            string? initial = commentElement.Element("Initial")?.Value;
            string? dateString = commentElement.Element("Date")?.Value;
            string? text = commentElement.Element("Text")?.Value;
            string? paraIndexString = commentElement.Element("ParagraphIndex")?.Value;

            // Validate required fields.
            if (author == null || initial == null || dateString == null || text == null || paraIndexString == null)
                continue; // Skip malformed entries.

            if (!DateTime.TryParse(dateString, out DateTime commentDate))
                commentDate = DateTime.Now;

            if (!int.TryParse(paraIndexString, out int paragraphIndex))
                continue; // Invalid index.

            // Ensure the target paragraph exists.
            if (doc.FirstSection?.Body?.Paragraphs == null ||
                paragraphIndex < 0 ||
                paragraphIndex >= doc.FirstSection.Body.Paragraphs.Count)
                continue; // Skip if index is out of range.

            // Create a new Comment node.
            Comment comment = new Comment(doc)
            {
                Author = author,
                Initial = initial,
                DateTime = commentDate
            };

            // The comment must contain at least one paragraph and run with visible text.
            Paragraph commentParagraph = new Paragraph(doc);
            Run commentRun = new Run(doc, text);
            commentParagraph.AppendChild(commentRun);
            comment.AppendChild(commentParagraph);

            // Attach the comment to the target paragraph.
            Paragraph targetParagraph = doc.FirstSection.Body.Paragraphs[paragraphIndex];
            // Adding the comment as a child of the paragraph anchors it to that paragraph.
            targetParagraph.AppendChild(comment);
        }

        // Save the resulting document with the imported comments.
        const string outputFileName = "DocumentWithImportedComments.docx";
        doc.Save(outputFileName);

        // Optional: Enumerate and display the imported comments to verify.
        var importedComments = doc.GetChildNodes(NodeType.Comment, true)
                                 .OfType<Comment>()
                                 .ToList();

        foreach (Comment c in importedComments)
        {
            Console.WriteLine($"Author: {c.Author}, Date: {c.DateTime:u}, Text: {c.GetText().Trim()}");
        }

        // Clean up temporary XML file.
        if (File.Exists(xmlFileName))
        {
            File.Delete(xmlFileName);
        }
    }
}
