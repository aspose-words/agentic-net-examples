using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;

public class Program
{
    public static void Main(string[] args)
    {
        // Prepare output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a sample DOCX with comments
        string samplePath = Path.Combine(outputDir, "sample.docx");
        CreateSampleDocument(samplePath);

        // Load the document
        Document doc = new Document(samplePath);

        // Enumerate all comments in the document
        List<Comment> comments = doc.GetChildNodes(NodeType.Comment, true)
                                   .OfType<Comment>()
                                   .ToList();

        // Export comments to CSV
        string csvPath = Path.Combine(outputDir, "comments.csv");
        ExportCommentsToCsv(comments, csvPath);
    }

    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with a comment
        builder.Writeln("This is the first paragraph.");
        Comment comment1 = new Comment(doc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };
        comment1.AppendChild(new Paragraph(doc));
        comment1.FirstParagraph?.AppendChild(new Run(doc, "First comment text."));
        // Attach the comment to the first paragraph if it exists
        Paragraph? firstPara = doc.FirstSection?.Body?.FirstParagraph;
        firstPara?.AppendChild(comment1);

        // Second paragraph with another comment
        builder.Writeln("This is the second paragraph.");
        Comment comment2 = new Comment(doc)
        {
            Author = "Jane Smith",
            Initial = "JS",
            DateTime = DateTime.Now.AddMinutes(-5)
        };
        comment2.AppendChild(new Paragraph(doc));
        comment2.FirstParagraph?.AppendChild(new Run(doc, "Second comment text, with a comma."));
        // Attach the comment to the last paragraph if it exists
        Paragraph? lastPara = doc.FirstSection?.Body?.LastParagraph;
        lastPara?.AppendChild(comment2);

        // Save the sample document
        doc.Save(filePath);
    }

    private static void ExportCommentsToCsv(IEnumerable<Comment> comments, string csvFilePath)
    {
        using (StreamWriter writer = new StreamWriter(csvFilePath, false))
        {
            // Write CSV header
            writer.WriteLine("Author,Date,Text");

            foreach (Comment comment in comments)
            {
                string author = EscapeForCsv(comment.Author ?? string.Empty);
                string date = comment.DateTime.ToString("o"); // ISO 8601 format
                string text = EscapeForCsv(comment.GetText().Trim());

                writer.WriteLine($"{author},{date},{text}");
            }
        }
    }

    private static string EscapeForCsv(string field)
    {
        if (field.Contains('"'))
        {
            field = field.Replace("\"", "\"\"");
        }

        if (field.Contains(',') || field.Contains('"') || field.Contains('\n') || field.Contains('\r'))
        {
            field = $"\"{field}\"";
        }

        return field;
    }
}
