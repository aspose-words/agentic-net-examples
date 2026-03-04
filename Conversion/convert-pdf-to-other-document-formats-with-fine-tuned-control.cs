using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfConversionExample
{
    static void Main()
    {
        // Create a new blank document and add some content.
        Document doc = new Document();                                   // create
        DocumentBuilder builder = new DocumentBuilder(doc);              // create
        builder.Writeln("This is a sample document that will be saved as PDF, then converted.");

        // Save the document as PDF with fine‑tuned options (e.g., outline levels).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Include only headings up to level 3 in the PDF outline.
            OutlineOptions = { HeadingsOutlineLevels = 3, ExpandedOutlineLevels = 1 }
        };
        doc.Save("Sample.pdf", pdfOptions);                              // save

        // Load the previously saved PDF back into an Aspose.Words Document.
        Document pdfDoc = new Document("Sample.pdf");                     // load

        // Convert the PDF to DOCX format.
        pdfDoc.Save("SampleConverted.docx", SaveFormat.Docx);            // save

        // Convert the same PDF to HTML format.
        pdfDoc.Save("SampleConverted.html", SaveFormat.Html);            // save

        // Convert the same PDF to plain text format.
        pdfDoc.Save("SampleConverted.txt", SaveFormat.Text);             // save

        Console.WriteLine("Conversion completed successfully.");
    }
}
