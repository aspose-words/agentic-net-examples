using System;
using System.Drawing;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add a paragraph with some text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This paragraph has a light gray background shading.");

        // Retrieve the paragraph that was just added.
        Paragraph paragraph = doc.FirstSection.Body.LastParagraph;

        // Set the paragraph shading background color to light gray.
        paragraph.ParagraphFormat.Shading.BackgroundPatternColor = Color.LightGray;

        // Save the document to a file.
        doc.Save("ParagraphShading.docx");
    }
}
