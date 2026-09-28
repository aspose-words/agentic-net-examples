using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Prepare a temporary folder for sample documents.
        string tempFolder = Path.Combine(Directory.GetCurrentDirectory(), "TempDocs");
        Directory.CreateDirectory(tempFolder);

        // Create sample documents with comments.
        CreateSampleDocument(Path.Combine(tempFolder, "Doc1.docx"), "First document", "Alice", "AL");
        CreateSampleDocument(Path.Combine(tempFolder, "Doc2.docx"), "Second document", "Bob", "BO");
        CreateSampleDocument(Path.Combine(tempFolder, "Doc3.docx"), "Third document", "Charlie", "CH");

        // Aggregate comments from all documents in the folder.
        var aggregatedComments = new List<(string SourceFile, Comment Comment)>();

        foreach (string filePath in Directory.GetFiles(tempFolder, "*.docx"))
        {
            Document doc = new Document(filePath);
            var comments = doc.GetChildNodes(NodeType.Comment, true)
                              .OfType<Comment>()
                              .ToList();

            foreach (Comment c in comments)
            {
                aggregatedComments.Add((Path.GetFileName(filePath), c));
            }
        }

        // Create a summary report document.
        Document reportDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(reportDoc);

        builder.Writeln("Comments Summary Report");
        builder.Writeln($"Generated on: {DateTime.Now:O}");
        builder.Writeln();

        foreach (var entry in aggregatedComments)
        {
            Comment c = entry.Comment;
            builder.Writeln($"Source Document: {entry.SourceFile}");
            builder.Writeln($"Author: {c.Author}");
            builder.Writeln($"Date: {c.DateTime:O}");
            builder.Writeln($"Text: {c.GetText().Trim()}");
            builder.Writeln();
        }

        // Save the report.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "CommentsReport.docx");
        reportDoc.Save(reportPath, SaveFormat.Docx);
    }

    // Helper method to create a document with a single comment.
    private static void CreateSampleDocument(string filePath, string paragraphText, string author, string initials)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a paragraph that will hold the comment.
        builder.Writeln(paragraphText);

        // Create a comment node.
        Comment comment = new Comment(doc)
        {
            Author = author,
            Initial = initials,
            DateTime = DateTime.Now
        };

        // Add visible content to the comment.
        Paragraph commentParagraph = new Paragraph(doc);
        commentParagraph.AppendChild(new Run(doc, $"Comment by {author} on {paragraphText}."));
        comment.AppendChild(commentParagraph);

        // Attach the comment to the last paragraph of the document.
        Paragraph? targetParagraph = doc.FirstSection?.Body?.LastParagraph;
        if (targetParagraph != null)
        {
            targetParagraph.AppendChild(comment);
        }

        // Save the document.
        doc.Save(filePath, SaveFormat.Docx);
    }
}
