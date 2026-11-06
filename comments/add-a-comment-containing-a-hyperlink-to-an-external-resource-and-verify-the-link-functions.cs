using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a paragraph that will hold the comment.
        builder.Writeln("This paragraph will have a comment with a hyperlink.");

        // Create a comment node with metadata.
        Comment comment = new Comment(doc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };

        // Add a paragraph inside the comment.
        Paragraph commentParagraph = new Paragraph(doc);
        comment.AppendChild(commentParagraph);

        // Add introductory text before the hyperlink.
        commentParagraph.AppendChild(new Run(doc, "For more information visit "));

        // Insert a hyperlink field inside the comment using a temporary builder.
        DocumentBuilder commentBuilder = new DocumentBuilder(doc);
        commentBuilder.MoveTo(commentParagraph);
        commentBuilder.InsertHyperlink("https://example.com", "example.com", false);

        // Append the comment to the first paragraph of the document body.
        Paragraph? firstParagraph = doc.FirstSection?.Body?.FirstParagraph;
        if (firstParagraph == null)
            throw new InvalidOperationException("The document does not contain a first paragraph.");

        firstParagraph.AppendChild(comment);

        // Save the document as PDF.
        string pdfPath = "CommentWithHyperlink.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Load the saved PDF back into Aspose.Words to verify the hyperlink text is present.
        Document pdfDoc = new Document(pdfPath);
        string pdfText = pdfDoc.GetText();

        bool containsLink = pdfText.Contains("example.com");
        Console.WriteLine(containsLink
            ? "Hyperlink text was found in the PDF."
            : "Hyperlink text was NOT found in the PDF.");
    }
}
