using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // -------------------------------------------------
        // Create first version of the document with two comments.
        // -------------------------------------------------
        Document doc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(doc1);

        // First paragraph with a comment from Alice.
        builder1.Writeln("First paragraph.");
        Comment comment1 = new Comment(doc1)
        {
            Author = "Alice",
            Initial = "A",
            DateTime = DateTime.Now
        };
        comment1.AppendChild(new Paragraph(doc1));
        comment1.FirstParagraph?.AppendChild(new Run(doc1, "First comment"));
        // Attach the comment to the first paragraph.
        Paragraph? firstPara = doc1.FirstSection?.Body?.FirstParagraph;
        firstPara?.AppendChild(comment1);

        // Second paragraph with a comment from Bob.
        builder1.Writeln("Second paragraph.");
        Comment comment2 = new Comment(doc1)
        {
            Author = "Bob",
            Initial = "B",
            DateTime = DateTime.Now
        };
        comment2.AppendChild(new Paragraph(doc1));
        comment2.FirstParagraph?.AppendChild(new Run(doc1, "Second comment"));
        // Attach the comment to the last paragraph (the one just added).
        Paragraph? lastPara = doc1.FirstSection?.Body?.LastParagraph;
        lastPara?.AppendChild(comment2);

        string doc1Path = Path.Combine(outputDir, "doc1.docx");
        doc1.Save(doc1Path);

        // -------------------------------------------------
        // Create second version: modify first comment, delete second, add new comment.
        // -------------------------------------------------
        Document doc2 = new Document(doc1Path);
        DocumentBuilder builder2 = new DocumentBuilder(doc2);

        // Modify first comment text (author Alice).
        Comment? firstComment = doc2.GetChildNodes(NodeType.Comment, true)
                                   .OfType<Comment>()
                                   .FirstOrDefault(c => c.Author == "Alice");
        if (firstComment?.FirstParagraph != null)
        {
            firstComment.FirstParagraph.RemoveAllChildren();
            firstComment.FirstParagraph.AppendChild(new Run(doc2, "First comment modified"));
        }

        // Delete second comment (author Bob).
        Comment? secondComment = doc2.GetChildNodes(NodeType.Comment, true)
                                    .OfType<Comment>()
                                    .FirstOrDefault(c => c.Author == "Bob");
        secondComment?.Remove();

        // Add a new comment by Charlie after a new paragraph.
        builder2.Writeln("Third paragraph.");
        Comment comment3 = new Comment(doc2)
        {
            Author = "Charlie",
            Initial = "C",
            DateTime = DateTime.Now
        };
        comment3.AppendChild(new Paragraph(doc2));
        comment3.FirstParagraph?.AppendChild(new Run(doc2, "Third comment"));
        Paragraph? newLastPara = doc2.FirstSection?.Body?.LastParagraph;
        newLastPara?.AppendChild(comment3);

        string doc2Path = Path.Combine(outputDir, "doc2.docx");
        doc2.Save(doc2Path);

        // -------------------------------------------------
        // Load both documents for comparison.
        // -------------------------------------------------
        Document oldDoc = new Document(doc1Path);
        Document newDoc = new Document(doc2Path);

        List<Comment> oldComments = oldDoc.GetChildNodes(NodeType.Comment, true)
                                          .OfType<Comment>()
                                          .ToList();

        List<Comment> newComments = newDoc.GetChildNodes(NodeType.Comment, true)
                                          .OfType<Comment>()
                                          .ToList();

        // Determine added comments.
        var added = newComments
            .Where(nc => !oldComments.Any(oc => oc.Author == nc.Author && oc.GetText().Trim() == nc.GetText().Trim()))
            .ToList();

        // Determine deleted comments.
        var deleted = oldComments
            .Where(oc => !newComments.Any(nc => nc.Author == oc.Author && nc.GetText().Trim() == oc.GetText().Trim()))
            .ToList();

        // Determine modified comments (same author, different text, matched by position).
        var modified = new List<(Comment Old, Comment New)>();
        int minCount = Math.Min(oldComments.Count, newComments.Count);
        for (int i = 0; i < minCount; i++)
        {
            Comment oc = oldComments[i];
            Comment nc = newComments[i];
            if (oc.Author == nc.Author && oc.GetText().Trim() != nc.GetText().Trim())
            {
                modified.Add((oc, nc));
            }
        }

        // -------------------------------------------------
        // Output results to console.
        // -------------------------------------------------
        Console.WriteLine("Added Comments:");
        foreach (Comment c in added)
        {
            Console.WriteLine($"- Author: {c.Author}, Text: {c.GetText().Trim()}");
        }

        Console.WriteLine("\nDeleted Comments:");
        foreach (Comment c in deleted)
        {
            Console.WriteLine($"- Author: {c.Author}, Text: {c.GetText().Trim()}");
        }

        Console.WriteLine("\nModified Comments:");
        foreach (var pair in modified)
        {
            Console.WriteLine($"- Author: {pair.Old.Author}");
            Console.WriteLine($"  Old Text: {pair.Old.GetText().Trim()}");
            Console.WriteLine($"  New Text: {pair.New.GetText().Trim()}");
        }

        // -------------------------------------------------
        // Save a simple report document summarizing the comparison.
        // -------------------------------------------------
        Document report = new Document();
        DocumentBuilder reportBuilder = new DocumentBuilder(report);

        reportBuilder.Writeln("Comments Comparison Report");
        reportBuilder.Writeln();

        reportBuilder.Writeln("Added Comments:");
        foreach (Comment c in added)
        {
            reportBuilder.Writeln($"Author: {c.Author}");
            reportBuilder.Writeln($"Text: {c.GetText().Trim()}");
            reportBuilder.Writeln();
        }

        reportBuilder.Writeln("Deleted Comments:");
        foreach (Comment c in deleted)
        {
            reportBuilder.Writeln($"Author: {c.Author}");
            reportBuilder.Writeln($"Text: {c.GetText().Trim()}");
            reportBuilder.Writeln();
        }

        reportBuilder.Writeln("Modified Comments:");
        foreach (var pair in modified)
        {
            reportBuilder.Writeln($"Author: {pair.Old.Author}");
            reportBuilder.Writeln($"Old Text: {pair.Old.GetText().Trim()}");
            reportBuilder.Writeln($"New Text: {pair.New.GetText().Trim()}");
            reportBuilder.Writeln();
        }

        string reportPath = Path.Combine(outputDir, "CommentsComparisonReport.docx");
        report.Save(reportPath);
    }
}
