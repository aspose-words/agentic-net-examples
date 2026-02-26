using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class AddBidiMarksExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a left‑to‑right paragraph.
        builder.Writeln("Hello world!"); // Default Bidi = false.

        // Insert a right‑to‑left paragraph.
        // Set the paragraph format to RTL before writing the text.
        builder.ParagraphFormat.Bidi = true;
        builder.Writeln("שלום עולם!"); // Hebrew text.

        // Insert another RTL paragraph (Arabic example).
        builder.Writeln("مرحبا بالعالم!"); // Arabic text.

        // Save the document as DOCX – the Bidi flag is stored in the paragraph properties.
        doc.Save("BidiParagraphs.docx");

        // OPTIONAL: Export the same document to plain text with explicit BiDi marks.
        // The TxtSaveOptions.AddBidiMarks property inserts U+200F before each RTL run.
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            Encoding = System.Text.Encoding.Unicode,
            AddBidiMarks = true
        };
        doc.Save("BidiParagraphs.txt", txtOptions);
    }
}
