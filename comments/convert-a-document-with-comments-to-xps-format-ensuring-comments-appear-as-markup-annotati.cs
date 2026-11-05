using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some content.
        builder.Writeln("This is the first paragraph.");
        builder.Writeln("This paragraph will have a comment attached.");

        // Retrieve the paragraph that will hold the comment.
        Paragraph? targetParagraph = doc.FirstSection?.Body?.LastParagraph;
        if (targetParagraph == null)
        {
            Console.WriteLine("Unable to locate the target paragraph.");
            return;
        }

        // Create a comment node with metadata.
        Comment comment = new Comment(doc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };

        // Add visible text to the comment.
        Paragraph commentParagraph = new Paragraph(doc);
        commentParagraph.AppendChild(new Run(doc, "Please review this paragraph."));
        comment.AppendChild(commentParagraph);

        // Attach the comment to the target paragraph.
        targetParagraph.AppendChild(comment);

        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document as DOCX (optional verification).
        string docxPath = Path.Combine(outputDir, "CommentedDocument.docx");
        doc.Save(docxPath, SaveFormat.Docx);

        // Save the document as XPS. Comments are rendered as markup annotations.
        string xpsPath = Path.Combine(outputDir, "CommentedDocument.xps");
        doc.Save(xpsPath, SaveFormat.Xps);

        // Indicate that processing is complete.
        Console.WriteLine($"DOCX saved to: {docxPath}");
        Console.WriteLine($"XPS saved to: {xpsPath}");
    }
}
