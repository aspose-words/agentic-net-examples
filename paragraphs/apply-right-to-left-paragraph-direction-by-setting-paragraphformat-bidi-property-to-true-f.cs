using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a regular English paragraph (optional, just for context).
        builder.Writeln("This is an English paragraph.");

        // Add an Arabic paragraph.
        builder.Writeln("مرحبا بالعالم");

        // The builder's CurrentParagraph now points to the Arabic paragraph we just added.
        // Set the paragraph direction to right‑to‑left.
        builder.CurrentParagraph.ParagraphFormat.Bidi = true;

        // Save the document to a file.
        doc.Save("Output.docx");
    }
}
