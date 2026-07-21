using System;
using System.Drawing;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the paragraph shading background color to light gray.
        builder.ParagraphFormat.Shading.BackgroundPatternColor = Color.LightGray;

        // Write some sample text to the paragraph.
        builder.Writeln("This paragraph has a light gray background shading.");

        // Save the document to the local file system.
        doc.Save("ParagraphShading.docx");
    }
}
