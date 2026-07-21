using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample XML file that represents exported comments.
        // -----------------------------------------------------------------
        string xmlContent =
@"<Comments>
    <Comment>
        <Author>Alice</Author>
        <Initial>AL</Initial>
        <DateTime>2023-01-15T10:30:00</DateTime>
        <Text>This is a comment for the first paragraph.</Text>
        <ParagraphIndex>0</ParagraphIndex>
    </Comment>
    <Comment>
        <Author>Bob</Author>
        <Initial>BO</Initial>
        <DateTime>2023-01-16T14:45:00</DateTime>
        <Text>Second paragraph needs review.</Text>
        <ParagraphIndex>1</ParagraphIndex>
    </Comment>
    <Comment>
        <Author>Charlie</Author>
        <Initial>CH</Initial>
        <DateTime>2023-01-17T09:15:00</DateTime>
        <Text>Final thoughts here.</Text>
        <ParagraphIndex>2</ParagraphIndex>
    </Comment>
</Comments>";
        string xmlPath = Path.Combine(outputDir, "comments.xml");
        File.WriteAllText(xmlPath, xmlContent);

        // -----------------------------------------------------------------
        // 2. Load the XML file.
        // -----------------------------------------------------------------
        XDocument xDoc = XDocument.Load(xmlPath);

        // -----------------------------------------------------------------
        // 3. Create a new Word document with three paragraphs.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Paragraph 1: Introduction.");
        builder.Writeln("Paragraph 2: Body content.");
        builder.Writeln("Paragraph 3: Conclusion.");

        // Retrieve all paragraphs in the document (in document order).
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true)
                            .OfType<Paragraph>()
                            .ToList();

        // -----------------------------------------------------------------
        // 4. Import comments from XML and attach them to the appropriate paragraph.
        // -----------------------------------------------------------------
        foreach (XElement commentElement in xDoc.Root.Elements("Comment"))
        {
            // Extract comment data safely.
            string author = (string?)commentElement.Element("Author") ?? "Unknown";
            string initial = (string?)commentElement.Element("Initial") ?? "";
            string text = (string?)commentElement.Element("Text") ?? "";
            string dateTimeStr = (string?)commentElement.Element("DateTime") ?? "";
            string paraIdxStr = (string?)commentElement.Element("ParagraphIndex") ?? "0";

            // Parse DateTime; fallback to DateTime.Now if parsing fails.
            DateTime dateTime = DateTime.TryParse(dateTimeStr, out DateTime dt) ? dt : DateTime.Now;

            // Parse paragraph index; ensure it is within range.
            int paraIdx = int.TryParse(paraIdxStr, out int idx) ? idx : 0;
            if (paraIdx < 0 || paraIdx >= paragraphs.Count)
                continue; // Skip invalid indices.

            Paragraph targetParagraph = paragraphs[paraIdx];

            // Create a new comment and set its properties.
            Comment comment = new Comment(doc, author, initial, dateTime);
            comment.SetText(text);

            // Attach the comment to the target paragraph.
            targetParagraph.AppendChild(comment);
        }

        // -----------------------------------------------------------------
        // 5. Save the resulting document.
        // -----------------------------------------------------------------
        string resultPath = Path.Combine(outputDir, "DocumentWithComments.docx");
        doc.Save(resultPath);

        // -----------------------------------------------------------------
        // 6. Optional: Enumerate and display the imported comments.
        // -----------------------------------------------------------------
        var importedComments = doc.GetChildNodes(NodeType.Comment, true)
                                 .OfType<Comment>()
                                 .ToList();

        foreach (Comment c in importedComments)
        {
            Console.WriteLine($"Author: {c.Author}, Date: {c.DateTime:u}, Text: {c.GetText().Trim()}");
        }
    }
}
