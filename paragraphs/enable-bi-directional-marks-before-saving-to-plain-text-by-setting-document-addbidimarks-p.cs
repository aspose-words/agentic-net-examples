using System;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to add paragraphs with mixed LTR and RTL text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, world!");   // LTR text.
        builder.Writeln("שלום עולם!");      // RTL (Hebrew) text.

        // Configure plain‑text saving options to insert Bi‑directional marks.
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            AddBidiMarks = true
        };

        // Save the document as plain text using the configured options.
        doc.Save("output.txt", saveOptions);
    }
}
