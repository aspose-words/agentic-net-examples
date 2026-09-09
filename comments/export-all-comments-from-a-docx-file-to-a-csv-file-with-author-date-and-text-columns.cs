using System;
using System.IO;
using System.Linq;
using Aspose.Words;

namespace ExportCommentsToCsv
{
    public class Program
    {
        public static void Main()
        {
            // Input and output file names.
            const string sampleDocPath = "sample.docx";
            const string csvPath = "comments.csv";

            // -----------------------------------------------------------------
            // 1. Create a sample DOCX document with a few comments.
            // -----------------------------------------------------------------
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // First paragraph with a comment.
            builder.Writeln("This is the first paragraph.");
            Comment comment1 = new Comment(doc, "Alice", "A", DateTime.Now);
            comment1.SetText("Review the wording of this paragraph.");
            builder.CurrentParagraph.AppendChild(comment1);

            // Second paragraph with a comment.
            builder.Writeln("Second paragraph follows.");
            Comment comment2 = new Comment(doc, "Bob", "B", DateTime.Now.AddMinutes(-15));
            comment2.SetText("Consider adding an example here.");
            builder.CurrentParagraph.AppendChild(comment2);

            // Save the sample document.
            doc.Save(sampleDocPath);

            // -----------------------------------------------------------------
            // 2. Load the document (simulating a real input file).
            // -----------------------------------------------------------------
            Document loadedDoc = new Document(sampleDocPath);

            // -----------------------------------------------------------------
            // 3. Enumerate all comments in the document.
            // -----------------------------------------------------------------
            var comments = loadedDoc
                .GetChildNodes(NodeType.Comment, true)
                .OfType<Comment>()
                .ToList();

            // -----------------------------------------------------------------
            // 4. Export comments to a CSV file with columns: Author, Date, Text.
            // -----------------------------------------------------------------
            using (var writer = new StreamWriter(csvPath))
            {
                // Write CSV header.
                writer.WriteLine("Author,Date,Text");

                foreach (Comment c in comments)
                {
                    string author = EscapeCsv(c.Author);
                    // ISO 8601 format for the date.
                    string date = EscapeCsv(c.DateTime.ToString("o"));
                    // Plain text of the comment.
                    string text = EscapeCsv(c.GetText().Trim());

                    writer.WriteLine($"{author},{date},{text}");
                }
            }

            Console.WriteLine($"Exported {comments.Count} comment(s) to '{csvPath}'.");
        }

        // Helper method to escape a CSV field according to RFC 4180.
        private static string EscapeCsv(string field)
        {
            if (field == null)
                return string.Empty;

            bool mustQuote = field.Contains(',') || field.Contains('"') || field.Contains('\r') || field.Contains('\n');
            if (mustQuote)
            {
                string escaped = field.Replace("\"", "\"\"");
                return $"\"{escaped}\"";
            }

            return field;
        }
    }
}
