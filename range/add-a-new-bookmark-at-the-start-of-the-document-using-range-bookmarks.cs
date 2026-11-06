using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a paragraph with some text to the document body.
        Paragraph paragraph = new Paragraph(doc);
        paragraph.AppendChild(new Run(doc, "Hello Aspose.Words!"));
        doc.FirstSection.Body.AppendChild(paragraph);

        // Insert a bookmark at the start of the document using DocumentBuilder.
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Move the cursor to the beginning of the first paragraph.
        builder.MoveTo(paragraph.FirstChild);
        // Start and end the bookmark at the same position.
        builder.StartBookmark("StartBookmark");
        builder.EndBookmark("StartBookmark");

        // Save the document to a file.
        doc.Save("Output.docx");
    }
}
