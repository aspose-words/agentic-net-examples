using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class AddBiDiMarksExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Left‑to‑right paragraph.
        builder.Writeln("Hello world!");

        // Right‑to‑left paragraph – enable BiDi layout for this paragraph.
        builder.ParagraphFormat.Bidi = true;
        builder.Writeln("שלום עולם!");   // Hebrew
        builder.Writeln("مرحبا بالعالم!"); // Arabic

        // Configure TxtSaveOptions to insert BiDi marks when exporting to plain text.
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            Encoding = Encoding.Unicode, // Use Unicode encoding to preserve marks.
            AddBidiMarks = true          // Insert U+200F before each BiDi run.
        };

        // Save the document as a .txt file with BiDi marks.
        doc.Save("BiDiMarkedDocument.txt", saveOptions);
    }
}
