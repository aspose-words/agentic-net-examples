using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a temporary folder for sample documents.
        string tempFolder = Path.Combine(Directory.GetCurrentDirectory(), "TempDocs");
        Directory.CreateDirectory(tempFolder);

        // Generate sample Word documents with comments.
        CreateSampleDocument(Path.Combine(tempFolder, "Document1.docx"), "First document text.", "Alice", "AL");
        CreateSampleDocument(Path.Combine(tempFolder, "Document2.docx"), "Second document text.", "Bob", "BO");

        // Aggregate comments from all documents in the folder.
        var aggregatedComments = new List<(string SourceFile, Comment Comment)>();

        foreach (string filePath in Directory.GetFiles(tempFolder, "*.docx"))
        {
            Document doc = new Document(filePath);
            var comments = doc.GetChildNodes(NodeType.Comment, true)
                              .OfType<Comment>()
                              .ToList();

            foreach (Comment comment in comments)
            {
                aggregatedComments.Add((Path.GetFileName(filePath), comment));
            }
        }

        // Create a summary report document.
        Document reportDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(reportDoc);

        builder.Writeln("Comments Summary Report");
        builder.Writeln($"Generated on {DateTime.Now:O}");
        builder.Writeln();

        foreach (var entry in aggregatedComments)
        {
            string sourceFile = entry.SourceFile;
            Comment comment = entry.Comment;

            builder.Writeln($"Source File: {sourceFile}");
            builder.Writeln($"Author: {comment.Author}");
            builder.Writeln($"Date: {comment.DateTime:O}");
            builder.Writeln($"Text: {comment.GetText().Trim()}");
            builder.Writeln();
        }

        // Save the report.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "CommentsReport.docx");
        reportDoc.Save(reportPath);
    }

    // Helper method to create a document with a single comment.
    private static void CreateSampleDocument(string filePath, string bodyText, string author, string initials)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a paragraph that will hold the comment.
        builder.Writeln(bodyText);

        // Create a comment node.
        Comment comment = new Comment(doc)
        {
            Author = author,
            Initial = initials,
            DateTime = DateTime.Now
        };
        // Add a paragraph and run inside the comment to make the comment visible.
        Paragraph commentParagraph = new Paragraph(doc);
        commentParagraph.AppendChild(new Run(doc, $"Comment by {author}: Review this text."));
        comment.AppendChild(commentParagraph);

        // Attach the comment to the last paragraph of the document.
        Paragraph? targetParagraph = doc.FirstSection?.Body?.LastParagraph;
        if (targetParagraph != null)
        {
            targetParagraph.AppendChild(comment);
        }

        // Save the document.
        doc.Save(filePath);
    }
}
