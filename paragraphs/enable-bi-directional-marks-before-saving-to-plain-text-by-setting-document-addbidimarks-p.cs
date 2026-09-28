using System;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample paragraphs containing left‑to‑right and right‑to‑left text.
        builder.Writeln("This is an English paragraph.");
        builder.Writeln("هذا نص عربي لتجربة الاتجاه من اليمين إلى اليسار.");

        // Enable insertion of Unicode bi‑directional marks when saving to plain text.
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            AddBidiMarks = true
        };

        // Save the document as plain text with the specified options.
        doc.Save("Output.txt", saveOptions);
    }
}
