using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables; // Needed for NodeType enum

namespace ExportCommentsToCsv
{
    public class Program
    {
        public static void Main()
        {
            // Create a sample DOCX file with comments.
            const string docPath = "sample.docx";
            CreateSampleDocument(docPath);

            // Load the document.
            Document doc = new Document(docPath);

            // Prepare the CSV output file.
            const string csvPath = "comments.csv";
            using (var writer = new StreamWriter(csvPath))
            {
                // Write CSV header.
                writer.WriteLine("Author,Date,Text");

                // Enumerate all comment nodes in the document.
                var comments = doc.GetChildNodes(NodeType.Comment, true)
                                  .OfType<Comment>()
                                  .ToList();

                foreach (Comment comment in comments)
                {
                    // Safely retrieve author, date, and text.
                    string author = comment.Author ?? string.Empty;
                    string date = comment.DateTime != DateTime.MinValue
                        ? comment.DateTime.ToString("o") // ISO 8601 format
                        : string.Empty;
                    string text = comment.GetText()?.Trim() ?? string.Empty;

                    // Escape double quotes by doubling them and wrap each field in quotes.
                    string Escape(string s) => $"\"{s.Replace("\"", "\"\"")}\"";

                    writer.WriteLine($"{Escape(author)},{Escape(date)},{Escape(text)}");
                }
            }

            // Inform that the process is complete.
            Console.WriteLine($"Exported {new FileInfo(csvPath).Length} bytes to '{csvPath}'.");
        }

        private static void CreateSampleDocument(string filePath)
        {
            // Initialize a new document and builder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // First paragraph with a comment.
            builder.Writeln("This is the first paragraph.");
            Comment comment1 = new Comment(doc, "Alice", "A", DateTime.Now);
            comment1.SetText("Review the first paragraph.");
            builder.CurrentParagraph.AppendChild(comment1);
            builder.MoveTo(comment1.AppendChild(new Paragraph(doc)));
            builder.Write("First comment text.");

            // Second paragraph with a comment.
            builder.Writeln("This is the second paragraph.");
            Comment comment2 = new Comment(doc, "Bob", "B", DateTime.Now.AddMinutes(-5));
            comment2.SetText("Check the second paragraph.");
            builder.CurrentParagraph.AppendChild(comment2);
            builder.MoveTo(comment2.AppendChild(new Paragraph(doc)));
            builder.Write("Second comment text.");

            // Save the document.
            doc.Save(filePath);
        }
    }
}
