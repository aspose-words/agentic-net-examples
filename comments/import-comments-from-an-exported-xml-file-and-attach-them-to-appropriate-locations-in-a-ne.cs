using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary folder for the example files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "CommentImportExample");
        Directory.CreateDirectory(workDir);

        // Path to the XML file that contains exported comment data.
        string xmlPath = Path.Combine(workDir, "comments.xml");

        // -----------------------------------------------------------------
        // 1. Create a sample XML file that represents exported comments.
        //    Each comment stores the index of the paragraph it belongs to,
        //    the author, initials, date/time and the comment text.
        // -----------------------------------------------------------------
        XDocument sampleXml = new XDocument(
            new XElement("Comments",
                new XElement("Comment",
                    new XElement("ParagraphIndex", 0),
                    new XElement("Author", "John Doe"),
                    new XElement("Initial", "JD"),
                    new XElement("DateTime", "2023-01-01T10:00:00"),
                    new XElement("Text", "Review the introduction.")
                ),
                new XElement("Comment",
                    new XElement("ParagraphIndex", 1),
                    new XElement("Author", "Jane Smith"),
                    new XElement("Initial", "JS"),
                    new XElement("DateTime", "2023-01-02T11:30:00"),
                    new XElement("Text", "Consider rephrasing this sentence.")
                )
            )
        );
        sampleXml.Save(xmlPath);

        // -----------------------------------------------------------------
        // 2. Load the XML file and parse comment information.
        // -----------------------------------------------------------------
        XDocument loadedXml = XDocument.Load(xmlPath);
        List<CommentInfo> commentInfos = new List<CommentInfo>();

        foreach (XElement commentElem in loadedXml.Root?.Elements("Comment") ?? new List<XElement>())
        {
            // Parse paragraph index.
            int paragraphIndex = int.TryParse(commentElem.Element("ParagraphIndex")?.Value, out int idx) ? idx : -1;
            if (paragraphIndex < 0) continue; // Skip invalid entries.

            // Parse author, initial and text.
            string author = commentElem.Element("Author")?.Value ?? "Unknown";
            string initial = commentElem.Element("Initial")?.Value ?? "";
            string text = commentElem.Element("Text")?.Value ?? "";

            // Parse date/time; fallback to now if parsing fails.
            DateTime dateTime = DateTime.TryParse(commentElem.Element("DateTime")?.Value, out DateTime dt)
                ? dt
                : DateTime.Now;

            commentInfos.Add(new CommentInfo
            {
                ParagraphIndex = paragraphIndex,
                Author = author,
                Initial = initial,
                DateTime = dateTime,
                Text = text
            });
        }

        // -----------------------------------------------------------------
        // 3. Create a new Word document and add some paragraphs.
        //    Keep references to the created Paragraph nodes for later use.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        List<Paragraph> paragraphs = new List<Paragraph>();

        // Add three sample paragraphs.
        for (int i = 0; i < 3; i++)
        {
            builder.Writeln($"This is paragraph {i + 1}.");
            // The paragraph just created becomes the current paragraph of the builder.
            Paragraph? currentPara = builder.CurrentParagraph;
            if (currentPara != null)
                paragraphs.Add(currentPara);
        }

        // -----------------------------------------------------------------
        // 4. Attach the imported comments to the appropriate paragraphs.
        // -----------------------------------------------------------------
        foreach (CommentInfo info in commentInfos)
        {
            // Ensure the target paragraph index exists.
            if (info.ParagraphIndex >= paragraphs.Count) continue;

            Paragraph targetParagraph = paragraphs[info.ParagraphIndex];

            // Create a new comment node.
            Comment comment = new Comment(doc, info.Author, info.Initial, info.DateTime);
            comment.SetText(info.Text);

            // Append the comment to the target paragraph.
            targetParagraph.AppendChild(comment);
        }

        // -----------------------------------------------------------------
        // 5. Save the resulting document.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(workDir, "DocumentWithImportedComments.docx");
        doc.Save(outputPath);

        // Optional: Write a short confirmation to the console.
        Console.WriteLine($"Document saved to: {outputPath}");
    }

    // Simple DTO to hold comment data parsed from XML.
    private class CommentInfo
    {
        public int ParagraphIndex { get; set; }
        public string Author { get; set; } = "";
        public string Initial { get; set; } = "";
        public DateTime DateTime { get; set; }
        public string Text { get; set; } = "";
    }
}
