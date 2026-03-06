using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Replacing; // Added for FindReplaceOptions

class Program
{
    static void Main()
    {
        // Load the DOCX document from the file system.
        Document doc = new Document("input.docx");

        // Example FAQ processing: replace a placeholder with actual text.
        // This demonstrates a simple find‑and‑replace operation.
        doc.Range.Replace("{{FAQ}}", "Frequently Asked Questions:", new FindReplaceOptions());

        // Configure image save options to render the document as PNG.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);
        // Render the first page only (zero‑based index). Remove or modify PageSet to render other pages.
        pngOptions.PageSet = new PageSet(0);

        // Save the rendered page as a PNG image.
        doc.Save("output.png", pngOptions);
    }
}
