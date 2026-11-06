using System;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a paragraph containing mixed left‑to‑right and right‑to‑left text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world! שלום עולם!");

        // Configure plain‑text save options to add bi‑directional marks.
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            AddBidiMarks = true
        };

        // Save the document as plain text with the configured options.
        doc.Save("output.txt", saveOptions);
    }
}
