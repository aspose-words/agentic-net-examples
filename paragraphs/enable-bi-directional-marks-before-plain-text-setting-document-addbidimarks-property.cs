using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a left‑to‑right paragraph.
        builder.Writeln("Hello world!");

        // Add a right‑to‑left paragraph (Hebrew example).
        builder.ParagraphFormat.Bidi = true;
        builder.Writeln("שלום עולם!");

        // Add another right‑to‑left paragraph (Arabic example).
        builder.Writeln("مرحبا بالعالم!");

        // Configure TxtSaveOptions to include BiDi marks in the plain‑text output.
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            Encoding = System.Text.Encoding.Unicode, // Use Unicode encoding to preserve characters.
            AddBidiMarks = true                       // Enable insertion of BiDi marks.
        };

        // Save the document as a .txt file using the configured options.
        doc.Save("OutputWithBidiMarks.txt", saveOptions);
    }
}
