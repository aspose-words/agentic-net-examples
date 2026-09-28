using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Words.Markup;

public class ImportCommentsFromXml
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Step 1: Create a sample XML file that contains exported comments.
        // The XML schema is simple: each <Comment> element has attributes:
        // ParagraphIndex (zero‑based), Author, Initial, DateTime (ISO 8601), and the comment text as element value.
        string xmlPath = Path.Combine(outputDir, "comments.xml");
        var sampleXml = new XDocument(
            new XElement("Comments",
                new XElement("Comment",
                    new XAttribute("ParagraphIndex", 0),
                    new XAttribute("Author", "Alice"),
                    new XAttribute("Initial", "AL"),
                    new XAttribute("DateTime", DateTime.Now.AddMinutes(-30).ToString("o")),
                    "Please review the introduction."),
                new XElement("Comment",
                    new XAttribute("ParagraphIndex", 2),
                    new XAttribute("Author", "Bob"),
                    new XAttribute("Initial", "BO"),
                    new XAttribute("DateTime", DateTime.Now.AddMinutes(-10).ToString("o")),
                    "Consider adding more examples here.")
            )
        );
        sampleXml.Save(xmlPath);

        // Step 2: Create a new document that will receive the comments.
        Document newDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(newDoc);

        // Add three paragraphs to the document. The paragraph index will be used to attach comments.
        for (int i = 0; i < 3; i++)
        {
            builder.Writeln($"This is paragraph {i + 1}.");
        }

        // Retrieve all paragraphs for easy indexing.
        var paragraphs = newDoc.GetChildNodes(NodeType.Paragraph, true)
                               .OfType<Paragraph>()
                               .ToList();

        // Step 3: Load the XML file and import each comment.
        XDocument loadedXml = XDocument.Load(xmlPath);
        var commentElements = loadedXml.Root?.Elements("Comment") ?? Enumerable.Empty<XElement>();

        foreach (var elem in commentElements)
        {
            // Parse required attributes safely.
            int paragraphIndex = (int?)elem.Attribute("ParagraphIndex") ?? -1;
            string? author = (string?)elem.Attribute("Author");
            string? initial = (string?)elem.Attribute("Initial");
            string? dateTimeStr = (string?)elem.Attribute("DateTime");
            string commentText = elem.Value ?? string.Empty;

            // Validate data before proceeding.
            if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count ||
                string.IsNullOrEmpty(author) ||
                string.IsNullOrEmpty(initial) ||
                string.IsNullOrEmpty(dateTimeStr))
            {
                continue; // Skip malformed entries.
            }

            if (!DateTime.TryParse(dateTimeStr, null, System.Globalization.DateTimeStyles.RoundtripKind, out DateTime commentDate))
            {
                commentDate = DateTime.Now;
            }

            // Create the Comment node.
            Comment commentNode = new Comment(newDoc)
            {
                Author = author,
                Initial = initial,
                DateTime = commentDate
            };

            // A comment must contain at least one paragraph with a run.
            commentNode.AppendChild(new Paragraph(newDoc));
            commentNode.FirstParagraph?.AppendChild(new Run(newDoc, commentText));

            // Attach the comment to the target paragraph.
            Paragraph targetParagraph = paragraphs[paragraphIndex];
            targetParagraph.AppendChild(commentNode);
        }

        // Step 4: Save the resulting document.
        string resultPath = Path.Combine(outputDir, "DocumentWithComments.docx");
        newDoc.Save(resultPath);
    }
}
