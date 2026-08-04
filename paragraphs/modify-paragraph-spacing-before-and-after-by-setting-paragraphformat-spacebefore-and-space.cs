using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the amount of spacing (in points) before and after each paragraph.
        builder.ParagraphFormat.SpaceBefore = 12; // 12 points before the paragraph.
        builder.ParagraphFormat.SpaceAfter = 12;  // 12 points after the paragraph.

        // Insert sample paragraphs that will inherit the spacing settings.
        builder.Writeln("First paragraph with custom spacing.");
        builder.Writeln("Second paragraph with the same custom spacing.");

        // Save the document to the file system.
        doc.Save("ParagraphSpacing.docx");
    }
}
