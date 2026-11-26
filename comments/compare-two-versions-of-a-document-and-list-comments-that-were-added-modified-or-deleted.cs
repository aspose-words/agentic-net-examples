using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Words.Layout;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create first version of the document with two comments.
        Document docV1 = CreateDocumentVersion1();

        // Create second version based on the first, modify comments.
        Document docV2 = CreateDocumentVersion2(docV1);

        // Save both versions for inspection (optional).
        docV1.Save("Document_V1.docx");
        docV2.Save("Document_V2.docx");

        // Enumerate comments in each version.
        List<Comment> commentsV1 = docV1.GetChildNodes(NodeType.Comment, true)
                                        .OfType<Comment>()
                                        .ToList();

        List<Comment> commentsV2 = docV2.GetChildNodes(NodeType.Comment, true)
                                        .OfType<Comment>()
                                        .ToList();

        // Determine added, deleted, and modified comments.
        int maxCount = Math.Max(commentsV1.Count, commentsV2.Count);
        for (int i = 0; i < maxCount; i++)
        {
            bool existsInV1 = i < commentsV1.Count;
            bool existsInV2 = i < commentsV2.Count;

            if (!existsInV1 && existsInV2)
            {
                // Comment added in version 2.
                Comment added = commentsV2[i];
                Console.WriteLine($"[Added] Author: {added.Author}, Text: {added.GetText().Trim()}");
            }
            else if (existsInV1 && !existsInV2)
            {
                // Comment deleted in version 2.
                Comment deleted = commentsV1[i];
                Console.WriteLine($"[Deleted] Author: {deleted.Author}, Text: {deleted.GetText().Trim()}");
            }
            else if (existsInV1 && existsInV2)
            {
                // Possible modification.
                Comment c1 = commentsV1[i];
                Comment c2 = commentsV2[i];

                string text1 = c1.GetText().Trim();
                string text2 = c2.GetText().Trim();

                if (!string.Equals(text1, text2, StringComparison.Ordinal))
                {
                    Console.WriteLine($"[Modified] Author: {c2.Author}");
                    Console.WriteLine($"    Original: {text1}");
                    Console.WriteLine($"    Updated : {text2}");
                }
            }
        }
    }

    // Creates the first document version with two comments.
    private static Document CreateDocumentVersion1()
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with a comment.
        builder.Writeln("First paragraph.");
        Paragraph? para1 = doc.FirstSection?.Body?.LastParagraph;
        if (para1 != null)
            AddCommentToParagraph(doc, para1, "Alice", "AL", "First comment.");

        // Second paragraph with a comment.
        builder.Writeln("Second paragraph.");
        Paragraph? para2 = doc.FirstSection?.Body?.LastParagraph;
        if (para2 != null)
            AddCommentToParagraph(doc, para2, "Bob", "BO", "Second comment.");

        return doc;
    }

    // Creates the second document version: updates first comment, removes second, adds a new one.
    private static Document CreateDocumentVersion2(Document baseDoc)
    {
        // Clone the base document to keep version 1 unchanged.
        Document doc = (Document)baseDoc.Clone(true);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Update the first comment's text.
        Comment? firstComment = doc.GetChildNodes(NodeType.Comment, true)
                                 .OfType<Comment>()
                                 .FirstOrDefault();
        if (firstComment != null)
        {
            // Clear existing content.
            firstComment.RemoveAllChildren();
            // Add updated text.
            Paragraph p = new Paragraph(doc);
            p.AppendChild(new Run(doc, "First comment - updated."));
            firstComment.AppendChild(p);
        }

        // Delete the second comment (Bob's) if it exists.
        List<Comment> allComments = doc.GetChildNodes(NodeType.Comment, true)
                                      .OfType<Comment>()
                                      .ToList();
        if (allComments.Count > 1)
        {
            Comment secondComment = allComments[1];
            secondComment.Remove();
        }

        // Add a new paragraph with a new comment.
        builder.Writeln("Third paragraph.");
        Paragraph? para3 = doc.FirstSection?.Body?.LastParagraph;
        if (para3 != null)
            AddCommentToParagraph(doc, para3, "Charlie", "CH", "Third comment.");

        return doc;
    }

    // Helper method to add a comment to a given paragraph.
    private static void AddCommentToParagraph(Document doc, Paragraph paragraph, string author, string initials, string commentText)
    {
        Comment comment = new Comment(doc)
        {
            Author = author,
            Initial = initials,
            DateTime = DateTime.Now
        };

        // Comment must contain at least one paragraph with a run.
        Paragraph commentParagraph = new Paragraph(doc);
        commentParagraph.AppendChild(new Run(doc, commentText));
        comment.AppendChild(commentParagraph);

        // Attach the comment to the target paragraph.
        paragraph.AppendChild(comment);
    }
}
