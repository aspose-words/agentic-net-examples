using System;
using System.Drawing;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with some text.
        builder.Writeln("This paragraph has a light gray background shading.");

        // Retrieve the paragraph we just added.
        Paragraph paragraph = doc.LastSection.Body.LastParagraph;

        // Set the shading background color to light gray.
        paragraph.ParagraphFormat.Shading.BackgroundPatternColor = Color.LightGray;

        // Save the document to a file.
        doc.Save("ParagraphShading.docx");
    }
}
