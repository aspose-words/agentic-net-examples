using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // The EmphasisMark enumeration does not contain a "Dot" value.
        // The closest visual representation is a solid circle placed above the text.
        // Apply that emphasis mark to the upcoming run.
        builder.Font.EmphasisMark = EmphasisMark.OverSolidCircle; // East‑Asian dot‑like mark

        // Write the text that will display the emphasis mark.
        builder.Write("Text with East Asian emphasis mark");

        // Optional: clear formatting for subsequent text.
        builder.Font.ClearFormatting();

        // Save the document to disk.
        doc.Save("EmphasisMarkExample.docx");
    }
}
