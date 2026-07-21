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

        // Set the paragraph direction to right‑to‑left.
        builder.ParagraphFormat.Bidi = true;

        // Write Arabic text into the paragraph.
        builder.Writeln("مرحبا بالعالم!"); // "Hello world!" in Arabic

        // Save the document to a file.
        string outputPath = "ParagraphBidi.docx";
        doc.Save(outputPath);
    }
}
