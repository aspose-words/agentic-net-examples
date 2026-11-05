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
        // Ensure Aspose.Words license is not required for this example.
        // Create a temporary folder for input and output files.
        string workFolder = Path.Combine(Directory.GetCurrentDirectory(), "CommentImportDemo");
        Directory.CreateDirectory(workFolder);

        // 1. Create a sample XML file that represents exported comments.
        string xmlPath = Path.Combine(workFolder, "comments.xml");
        CreateSampleCommentsXml(xmlPath);

        // 2. Load comments from the XML file.
        var commentData = LoadCommentsFromXml(xmlPath);

        // 3. Create a new document with several paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add five sample paragraphs.
        for (int i = 0; i < 5; i++)
        {
            builder.Writeln($"This is paragraph {i + 1}.");
        }

        // 4. Attach comments to the appropriate paragraphs based on the XML data.
        foreach (var data in commentData)
        {
            // Validate paragraph index.
            if (data.ParagraphIndex < 0 || data.ParagraphIndex >= doc.FirstSection.Body.Paragraphs.Count)
                continue; // Skip invalid indices.

            // Retrieve the target paragraph.
            var targetParagraph = doc.FirstSection.Body.Paragraphs[data.ParagraphIndex];

            // Create a new comment.
            Comment comment = new Comment(doc)
            {
                Author = data.Author,
                Initial = data.Initial,
                DateTime = data.DateTime
            };

            // Add a paragraph and run inside the comment to hold the comment text.
            comment.AppendChild(new Paragraph(doc));
            comment.FirstParagraph?.AppendChild(new Run(doc, data.Text));

            // Append the comment to the target paragraph.
            targetParagraph.AppendChild(comment);
        }

        // 5. Save the resulting document.
        string outputPath = Path.Combine(workFolder, "DocumentWithComments.docx");
        doc.Save(outputPath, SaveFormat.Docx);

        // The example finishes here. No interactive prompts are used.
    }

    // Helper method to create a sample XML file with comment information.
    private static void CreateSampleCommentsXml(string filePath)
    {
        XDocument xDoc = new XDocument(
            new XElement("Comments",
                new XElement("Comment",
                    new XAttribute("Author", "Alice"),
                    new XAttribute("Initial", "AL"),
                    new XAttribute("DateTime", DateTime.Now.AddDays(-2).ToString("o")),
                    new XAttribute("ParagraphIndex", "1"),
                    new XElement("Text", "Please review this paragraph.")
                ),
                new XElement("Comment",
                    new XAttribute("Author", "Bob"),
                    new XAttribute("Initial", "BO"),
                    new XAttribute("DateTime", DateTime.Now.AddDays(-1).ToString("o")),
                    new XAttribute("ParagraphIndex", "3"),
                    new XElement("Text", "Consider adding more details here.")
                )
            )
        );

        xDoc.Save(filePath);
    }

    // Data structure to hold comment information loaded from XML.
    private class CommentInfo
    {
        public string Author { get; set; } = "";
        public string Initial { get; set; } = "";
        public DateTime DateTime { get; set; }
        public int ParagraphIndex { get; set; }
        public string Text { get; set; } = "";
    }

    // Helper method to parse the XML file into a list of CommentInfo objects.
    private static System.Collections.Generic.List<CommentInfo> LoadCommentsFromXml(string filePath)
    {
        var list = new System.Collections.Generic.List<CommentInfo>();

        XDocument xDoc = XDocument.Load(filePath);
        foreach (XElement elem in xDoc.Root?.Elements("Comment") ?? Enumerable.Empty<XElement>())
        {
            string author = (string?)elem.Attribute("Author") ?? "";
            string initial = (string?)elem.Attribute("Initial") ?? "";
            string dateStr = (string?)elem.Attribute("DateTime") ?? "";
            string indexStr = (string?)elem.Attribute("ParagraphIndex") ?? "0";
            string text = elem.Element("Text")?.Value ?? "";

            if (!int.TryParse(indexStr, out int paragraphIndex))
                paragraphIndex = 0;

            DateTime dateTime = DateTime.Now;
            if (!DateTime.TryParse(dateStr, null, System.Globalization.DateTimeStyles.RoundtripKind, out dateTime))
                dateTime = DateTime.Now;

            list.Add(new CommentInfo
            {
                Author = author,
                Initial = initial,
                DateTime = dateTime,
                ParagraphIndex = paragraphIndex,
                Text = text
            });
        }

        return list;
    }
}
