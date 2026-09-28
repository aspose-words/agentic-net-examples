using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new document and add a paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Original paragraph with comment.");

        // Create a comment and attach it to the first paragraph.
        Comment comment = new Comment(doc)
        {
            Author = "Tester",
            Initial = "TS",
            DateTime = DateTime.Now
        };
        // A comment must contain at least one paragraph.
        comment.AppendChild(new Paragraph(doc));
        // Add visible text to the comment.
        if (comment.FirstParagraph != null)
        {
            comment.FirstParagraph.AppendChild(new Run(doc, "Please review this paragraph."));
        }

        Paragraph? firstParagraph = doc.FirstSection?.Body?.FirstParagraph;
        if (firstParagraph != null)
        {
            firstParagraph.AppendChild(comment);
        }

        // Record comment Id and the index of its parent paragraph before insertion.
        int commentIdBefore = comment.Id;
        int paragraphIndexBefore = firstParagraph != null
            ? doc.FirstSection!.Body!.Paragraphs.IndexOf(firstParagraph)
            : -1;

        // Insert a new paragraph before the original paragraph.
        DocumentBuilder insertBuilder = new DocumentBuilder(doc);
        insertBuilder.MoveToDocumentStart();
        insertBuilder.Writeln("Inserted new paragraph before the original.");

        // Retrieve the comment after insertion.
        Comment? retrievedComment = doc.GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .FirstOrDefault();

        // Determine the paragraph that now contains the comment.
        Paragraph? commentParentParagraph = retrievedComment?.ParentNode as Paragraph;
        int paragraphIndexAfter = commentParentParagraph != null
            ? doc.FirstSection!.Body!.Paragraphs.IndexOf(commentParentParagraph)
            : -1;

        // Validate that the comment Id is unchanged and its paragraph index shifted by one.
        bool idUnchanged = retrievedComment != null && retrievedComment.Id == commentIdBefore;
        bool indexShifted = paragraphIndexAfter == paragraphIndexBefore + 1;

        // Output validation results.
        Console.WriteLine($"Comment Id unchanged: {idUnchanged}");
        Console.WriteLine($"Paragraph index before: {paragraphIndexBefore}, after: {paragraphIndexAfter}, shifted correctly: {indexShifted}");

        // Save the document for manual inspection if needed.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CommentReferenceUpdate.docx");
        doc.Save(outputPath);
    }
}
