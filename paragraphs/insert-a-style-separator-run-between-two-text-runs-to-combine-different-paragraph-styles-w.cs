using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First part of the paragraph uses the Heading1 style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Write("This is heading style");

        // Insert a style separator run. Text after this run can have a different style
        // while still belonging to the same paragraph.
        builder.InsertStyleSeparator();

        // Second part of the paragraph should appear with the normal (default) character style.
        // Use a character style because Normal is a paragraph style and cannot be applied to a run.
        builder.Font.StyleIdentifier = StyleIdentifier.DefaultParagraphFont;
        builder.Write(" and this is normal style.");

        // Save the document to disk.
        doc.Save("StyleSeparator.docx");
    }
}
