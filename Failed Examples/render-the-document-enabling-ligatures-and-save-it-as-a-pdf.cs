// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable standard ligatures for the current font.
        // This will cause character pairs like "fi" and "fl" to be rendered as ligatures.
        builder.Font.Ligature = FontLigature.Standard;

        // Add some text that contains characters which can form ligatures.
        builder.Writeln("Office filing");

        // Create PDF save options (default settings are sufficient for this task).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the document as a PDF file using the save options.
        doc.Save("Output.pdf", pdfOptions);
    }
}
