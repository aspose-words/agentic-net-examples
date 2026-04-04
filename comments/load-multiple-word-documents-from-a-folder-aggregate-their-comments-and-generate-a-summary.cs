using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a temporary working directory.
        string workDir = Path.Combine(Path.GetTempPath(), "AsposeCommentsDemo");
        Directory.CreateDirectory(workDir);

        // -----------------------------------------------------------------
        // Step 1: Generate sample Word documents with comments.
        // -----------------------------------------------------------------
        string[] sampleFiles = new string[2];
        for (int i = 0; i < sampleFiles.Length; i++)
        {
            string filePath = Path.Combine(workDir, $"SampleDocument{i + 1}.docx");
            CreateSampleDocument(filePath, $"Document {i + 1}", $"Author{i + 1}", $"A{i + 1}");
            sampleFiles[i] = filePath;
        }

        // -----------------------------------------------------------------
        // Step 2: Load each document, collect its comments.
        // -----------------------------------------------------------------
        var aggregatedComments = new System.Collections.Generic.List<(string SourceFile, Comment Comment)>();

        foreach (string file in Directory.GetFiles(workDir, "*.docx"))
        {
            Document doc = new Document(file);

            var comments = doc.GetChildNodes(NodeType.Comment, true)
                              .OfType<Comment>()
                              .ToList();

            foreach (Comment c in comments)
            {
                aggregatedComments.Add((Path.GetFileName(file), c));
            }
        }

        // -----------------------------------------------------------------
        // Step 3: Create a summary report document.
        // -----------------------------------------------------------------
        Document reportDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(reportDoc);

        builder.Writeln("Comments Summary Report");
        builder.Writeln($"Generated on: {DateTime.Now:O}");
        builder.Writeln();

        foreach (var entry in aggregatedComments)
        {
            builder.Writeln($"Source Document : {entry.SourceFile}");
            builder.Writeln($"Author          : {entry.Comment.Author}");
            builder.Writeln($"Date/Time       : {entry.Comment.DateTime:O}");
            builder.Writeln($"Comment Text    : {entry.Comment.GetText().Trim()}");
            builder.Writeln(); // Blank line between entries
        }

        // Save the report.
        string reportPath = Path.Combine(workDir, "CommentsSummaryReport.docx");
        reportDoc.Save(reportPath);

        // The example finishes without waiting for user input.
    }

    // Helper method to create a simple document with a single paragraph and a comment.
    private static void CreateSampleDocument(string filePath, string paragraphText, string author, string initials)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a paragraph that will host the comment.
        builder.Writeln(paragraphText);

        // Create a comment attached to the last paragraph.
        Comment comment = new Comment(doc)
        {
            Author = author,
            Initial = initials,
            DateTime = DateTime.Now
        };
        // Ensure the comment has visible content.
        comment.AppendChild(new Paragraph(doc));
        comment.FirstParagraph.AppendChild(new Run(doc, $"This is a comment from {author}."));

        // Append the comment to the paragraph.
        Paragraph lastParagraph = doc.FirstSection.Body.LastParagraph;
        lastParagraph?.AppendChild(comment);

        // Save the document.
        doc.Save(filePath);
    }
}
