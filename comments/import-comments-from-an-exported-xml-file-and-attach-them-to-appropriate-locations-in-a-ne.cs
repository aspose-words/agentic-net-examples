using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;

#nullable enable

public class Program
{
    public static void Main()
    {
        // 1. Create a sample XML file that contains exported comment data.
        const string xmlFileName = "comments.xml";

        var xmlContent =
@"<Comments>
    <Comment>
        <ParagraphIndex>0</ParagraphIndex>
        <Author>John Doe</Author>
        <Initial>JD</Initial>
        <Date>2023-01-01T10:00:00</Date>
        <Text>This is a comment for the first paragraph.</Text>
    </Comment>
    <Comment>
        <ParagraphIndex>2</ParagraphIndex>
        <Author>Jane Smith</Author>
        <Initial>JS</Initial>
        <Date>2023-02-15T14:30:00</Date>
        <Text>Second comment, attached to the third paragraph.</Text>
    </Comment>
</Comments>";

        File.WriteAllText(xmlFileName, xmlContent);

        // 2. Load the XML file and parse comment information.
        XDocument xDoc = XDocument.Load(xmlFileName);
        var commentElements = xDoc.Root?.Elements("Comment") ?? Enumerable.Empty<XElement>();

        // 3. Create a new Word document with a few paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Paragraph 1: Lorem ipsum dolor sit amet.");
        builder.Writeln("Paragraph 2: Consectetur adipiscing elit.");
        builder.Writeln("Paragraph 3: Sed do eiusmod tempor incididunt.");

        // 4. Attach comments from the XML to the appropriate paragraphs.
        foreach (var elem in commentElements)
        {
            // Parse required fields with safety checks.
            int? paragraphIndex = (int?)elem.Element("ParagraphIndex");
            string? author = (string?)elem.Element("Author");
            string? initial = (string?)elem.Element("Initial");
            string? dateString = (string?)elem.Element("Date");
            string? text = (string?)elem.Element("Text");

            // Validate mandatory data.
            if (paragraphIndex == null || author == null || initial == null || dateString == null || text == null)
                continue; // Skip malformed entries.

            // Ensure the paragraph index is within the document range.
            ParagraphCollection? paragraphs = doc.FirstSection?.Body?.Paragraphs;
            if (paragraphs == null || paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
                continue; // No such paragraph; skip.

            Paragraph targetParagraph = paragraphs[paragraphIndex.Value];

            // Parse the date.
            if (!DateTime.TryParse(dateString, out DateTime commentDate))
                commentDate = DateTime.Now;

            // Create the comment node with metadata.
            Comment comment = new Comment(doc, author, initial, commentDate);
            comment.SetText(text);

            // Append the comment to the target paragraph.
            targetParagraph.AppendChild(comment);
        }

        // 5. Save the resulting document.
        const string outputDoc = "DocumentWithComments.docx";
        doc.Save(outputDoc);

        // 6. Enumerate and display comment information to the console.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        foreach (Comment c in comments)
        {
            Console.WriteLine($"Author: {c.Author}, Date: {c.DateTime:u}, Text: {c.GetText().Trim()}");
        }

        // Clean up the temporary XML file (optional).
        if (File.Exists(xmlFileName))
            File.Delete(xmlFileName);
    }
}
