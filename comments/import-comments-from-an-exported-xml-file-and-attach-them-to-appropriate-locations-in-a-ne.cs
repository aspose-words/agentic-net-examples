using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary folder for all generated files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // -----------------------------------------------------------------
        // 1. Create a sample XML file that represents exported comments.
        // -----------------------------------------------------------------
        string xmlPath = Path.Combine(workDir, "comments.xml");
        XDocument commentXml = new XDocument(
            new XElement("Comments",
                new XElement("Comment",
                    new XElement("Author", "John Doe"),
                    new XElement("Initial", "JD"),
                    new XElement("Date", "2023-01-01T10:30:00"),
                    new XElement("Text", "Review this paragraph."),
                    new XElement("ParagraphIndex", "0")
                ),
                new XElement("Comment",
                    new XElement("Author", "Jane Smith"),
                    new XElement("Initial", "JS"),
                    new XElement("Date", "2023-01-02T14:45:00"),
                    new XElement("Text", "Consider rephrasing."),
                    new XElement("ParagraphIndex", "2")
                )
            )
        );
        commentXml.Save(xmlPath);

        // -----------------------------------------------------------------
        // 2. Load the XML file and parse comment data.
        // -----------------------------------------------------------------
        XDocument loadedXml = XDocument.Load(xmlPath);
        var commentElements = loadedXml.Root?.Elements("Comment") ?? Enumerable.Empty<XElement>();

        // -----------------------------------------------------------------
        // 3. Create a new Word document with some paragraphs.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Store references to the created paragraphs for later comment attachment.
        List<Paragraph> paragraphs = new List<Paragraph>();

        for (int i = 0; i < 3; i++)
        {
            builder.Writeln($"Paragraph {i + 1}");
            // After Writeln, CurrentParagraph points to the paragraph we just added.
            Paragraph? para = builder.CurrentParagraph;
            if (para != null)
                paragraphs.Add(para);
        }

        // -----------------------------------------------------------------
        // 4. Attach each imported comment to the appropriate paragraph.
        // -----------------------------------------------------------------
        foreach (XElement elem in commentElements)
        {
            string author = (string?)elem.Element("Author") ?? "Unknown";
            string initial = (string?)elem.Element("Initial") ?? "";
            string dateStr = (string?)elem.Element("Date") ?? DateTime.MinValue.ToString("o");
            string text = (string?)elem.Element("Text") ?? "";
            string indexStr = (string?)elem.Element("ParagraphIndex") ?? "0";

            // Parse date and paragraph index safely.
            DateTime dateTime = DateTime.TryParse(dateStr, out DateTime dt) ? dt : DateTime.Now;
            int paraIndex = int.TryParse(indexStr, out int idx) ? idx : 0;

            // Ensure the paragraph index is within the existing range.
            if (paraIndex < 0 || paraIndex >= paragraphs.Count)
                continue; // Skip invalid entries.

            Paragraph targetParagraph = paragraphs[paraIndex];

            // Create the comment and set its text.
            Comment comment = new Comment(doc, author, initial, dateTime);
            comment.SetText(text);

            // Append the comment to the target paragraph.
            targetParagraph.AppendChild(comment);
        }

        // -----------------------------------------------------------------
        // 5. Save the resulting document.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(workDir, "DocumentWithImportedComments.docx");
        doc.Save(outputPath);

        // The example finishes without waiting for user input.
    }
}
