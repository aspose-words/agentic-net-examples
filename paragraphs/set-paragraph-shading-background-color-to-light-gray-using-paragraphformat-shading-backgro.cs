using System;
using Aspose.Words;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the paragraph shading background color to light gray.
        builder.ParagraphFormat.Shading.BackgroundPatternColor = Color.LightGray;

        // Add a paragraph; the shading will be applied to this paragraph.
        builder.Writeln("This paragraph has a light gray background shading.");

        // Save the document to a file.
        doc.Save("ParagraphShading.docx");
    }
}
