using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply the built‑in "Quote" style to the upcoming paragraph.
        builder.ParagraphFormat.StyleName = "Quote";

        // Increase the left indent (in points) for emphasis.
        builder.ParagraphFormat.LeftIndent = 20.0;

        // Write the paragraph text.
        builder.Writeln("This paragraph uses the built‑in Quote style and has an increased left indent.");

        // Save the document to the local file system.
        doc.Save("QuoteStyle.docx");
    }
}
