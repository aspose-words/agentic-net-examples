using System;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;

public class ExportCommentsToCsv
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a sample DOCX file with comments.
        string docPath = Path.Combine(outputDir, "sample.docx");
        CreateSampleDocument(docPath);

        // Load the document.
        Document doc = new Document(docPath);

        // Enumerate all comments in the document.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        // Prepare CSV file.
        string csvPath = Path.Combine(outputDir, "comments.csv");
        using (var writer = new StreamWriter(csvPath, false, Encoding.UTF8))
        {
            // Write CSV header.
            writer.WriteLine("Author,Date,Text");

            // Write each comment as a CSV line.
            foreach (Comment comment in comments)
            {
                string author = comment.Author ?? string.Empty;
                string date = comment.DateTime.ToString("o"); // ISO 8601 format.
                string text = comment.GetText()?.Trim() ?? string.Empty;

                // Escape CSV fields.
                author = EscapeCsvField(author);
                date = EscapeCsvField(date);
                text = EscapeCsvField(text);

                writer.WriteLine($"{author},{date},{text}");
            }
        }

        // Example completed. Output files are in the 'output' folder.
    }

    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with a comment.
        builder.Writeln("This is the first paragraph.");
        Comment comment1 = new Comment(doc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };
        comment1.AppendChild(new Paragraph(doc));
        comment1.FirstParagraph?.AppendChild(new Run(doc, "Review the first paragraph."));
        // Append the comment to the paragraph that was just created.
        builder.CurrentParagraph?.AppendChild(comment1);

        // Second paragraph with another comment.
        builder.Writeln("This is the second paragraph.");
        Comment comment2 = new Comment(doc)
        {
            Author = "Jane Smith",
            Initial = "JS",
            DateTime = DateTime.Now.AddMinutes(-5)
        };
        comment2.AppendChild(new Paragraph(doc));
        comment2.FirstParagraph?.AppendChild(new Run(doc, "Check the data in this paragraph."));
        builder.CurrentParagraph?.AppendChild(comment2);

        // Save the document.
        doc.Save(filePath);
    }

    private static string EscapeCsvField(string field)
    {
        if (field.Contains('"') || field.Contains(',') || field.Contains('\n') || field.Contains('\r'))
        {
            string escaped = field.Replace("\"", "\"\"");
            return $"\"{escaped}\"";
        }
        return field;
    }
}
