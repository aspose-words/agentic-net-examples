using Aspose.Words;
using System;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample Arabic text.
        string arabicText = "مرحبا بالعالم!";

        // Apply right‑to‑left direction to the current paragraph.
        builder.ParagraphFormat.Bidi = true;

        // Write the Arabic text into the paragraph.
        builder.Writeln(arabicText);

        // Save the document to a file.
        doc.Save("ArabicParagraph.docx");
    }
}
