using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create an output folder for the generated files.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Initialize a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will hold a comment.
        builder.Writeln("This paragraph will have a comment attached to it.");

        // Create a comment node with author information and comment text.
        Comment comment = new Comment(doc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };
        // The comment must contain at least one paragraph and run to be visible.
        comment.AppendChild(new Paragraph(doc));
        comment.FirstParagraph?.AppendChild(new Run(doc, "Please review this paragraph."));

        // Attach the comment to the first paragraph of the document.
        Paragraph? firstParagraph = doc.FirstSection?.Body?.FirstParagraph;
        if (firstParagraph != null)
        {
            firstParagraph.AppendChild(comment);
        }

        // Save the document as DOCX (optional, for verification purposes).
        string docxPath = Path.Combine(outputDir, "CommentedDocument.docx");
        doc.Save(docxPath);

        // Save the document as XPS. Comments are rendered as markup annotations.
        string xpsPath = Path.Combine(outputDir, "CommentedDocument.xps");
        XpsSaveOptions xpsOptions = new XpsSaveOptions(SaveFormat.Xps);
        doc.Save(xpsPath, xpsOptions);
    }
}
