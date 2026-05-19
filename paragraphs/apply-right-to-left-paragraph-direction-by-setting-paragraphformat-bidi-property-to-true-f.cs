using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the paragraph direction to right‑to‑left.
        builder.ParagraphFormat.Bidi = true;

        // Write an Arabic sentence.
        builder.Writeln("مرحبا بالعالم!"); // "Hello world!" in Arabic

        // Save the document to the current directory.
        string outputPath = System.IO.Path.Combine(Environment.CurrentDirectory, "ArabicRtl.docx");
        doc.Save(outputPath);
    }
}
