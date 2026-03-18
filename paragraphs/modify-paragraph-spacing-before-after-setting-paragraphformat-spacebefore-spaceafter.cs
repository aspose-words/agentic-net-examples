using Aspose.Words;
using System;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set custom spacing before and after each paragraph (points).
        builder.ParagraphFormat.SpaceBefore = 12; // 12 points before the paragraph
        builder.ParagraphFormat.SpaceAfter = 12;  // 12 points after the paragraph

        // Disable automatic spacing so the custom values are applied.
        builder.ParagraphFormat.SpaceBeforeAuto = false;
        builder.ParagraphFormat.SpaceAfterAuto = false;

        // Insert sample paragraphs that will inherit the spacing settings.
        builder.Writeln("First paragraph with custom spacing.");
        builder.Writeln("Second paragraph with the same spacing.");

        // Save the document to a file.
        doc.Save("ParagraphSpacing.docx");
    }
}
