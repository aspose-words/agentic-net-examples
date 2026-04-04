using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;

namespace CommentImportExample
{
    public class Program
    {
        public static void Main()
        {
            // Prepare output folder.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // 1. Create a sample XML file that contains exported comments.
            string xmlPath = Path.Combine(outputDir, "comments.xml");
            CreateSampleCommentsXml(xmlPath);

            // 2. Create a new document with some paragraphs.
            Document doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("First paragraph of the document.");
            builder.Writeln("Second paragraph of the document.");
            builder.Writeln("Third paragraph of the document.");

            // 3. Load comments from the XML and attach them to the appropriate paragraphs.
            XDocument xDoc = XDocument.Load(xmlPath);
            var commentElements = xDoc.Root?.Elements("Comment") ?? Enumerable.Empty<XElement>();

            foreach (XElement elem in commentElements)
            {
                // Parse required data.
                int paragraphIndex = (int?)elem.Element("ParagraphIndex") ?? 0;
                string author = (string?)elem.Element("Author") ?? "Unknown";
                string initial = (string?)elem.Element("Initial") ?? "";
                DateTime dateTime = DateTime.TryParse((string?)elem.Element("DateTime"), out DateTime dt) ? dt : DateTime.Now;
                string text = (string?)elem.Element("Text") ?? "";

                // Ensure the paragraph index is within range.
                if (paragraphIndex < 0 || paragraphIndex >= doc.FirstSection.Body.Paragraphs.Count)
                    continue; // Skip invalid entries.

                // Retrieve the target paragraph.
                var targetParagraph = doc.FirstSection.Body.Paragraphs[paragraphIndex];

                // Create the comment.
                Comment comment = new Comment(doc, author, initial, dateTime);
                comment.SetText(text);

                // Attach the comment to the paragraph.
                targetParagraph.AppendChild(comment);
            }

            // 4. Save the resulting document.
            string resultPath = Path.Combine(outputDir, "DocumentWithImportedComments.docx");
            doc.Save(resultPath);

            // 5. (Optional) Enumerate and display the imported comments.
            var comments = doc.GetChildNodes(NodeType.Comment, true)
                              .OfType<Comment>()
                              .ToList();

            foreach (Comment c in comments)
            {
                Console.WriteLine($"Comment ID {c.Id} by {c.Author} on {c.DateTime:u}");
                Console.WriteLine($"Text: {c.GetText().Trim()}");
                Console.WriteLine();
            }
        }

        // Helper method to create a sample XML file with comment data.
        private static void CreateSampleCommentsXml(string filePath)
        {
            XDocument doc = new XDocument(
                new XElement("Comments",
                    new XElement("Comment",
                        new XElement("ParagraphIndex", 0),
                        new XElement("Author", "John Doe"),
                        new XElement("Initial", "JD"),
                        new XElement("DateTime", "2023-01-01T10:00:00"),
                        new XElement("Text", "Review the first paragraph.")
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

            doc.Save(filePath);
        }
    }
}
