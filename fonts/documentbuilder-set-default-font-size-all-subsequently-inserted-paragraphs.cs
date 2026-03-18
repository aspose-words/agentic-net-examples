using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document
        Document doc = new Document();

        // Set the document‑wide default font size.
        // All paragraphs that are inserted after this point will inherit this size
        // unless their own Font.Size is overridden.
        doc.Styles.DefaultFont.Size = 14; // 14 points

        // Attach a DocumentBuilder to the document
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert paragraphs – they will use the default font size set above
        builder.Writeln("First paragraph using the default font size.");
        builder.Writeln("Second paragraph also using the default font size.");

        // Save the document
        doc.Save("DefaultFontSize.docx");
    }
}
