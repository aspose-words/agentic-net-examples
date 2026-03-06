using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfAExample
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Explain when PDF/A should be used.
        builder.Writeln("When to use PDF/A:");
        builder.Writeln("- Archival storage of documents that must remain readable for decades.");
        builder.Writeln("- Legal, governmental, and financial records.");
        builder.Writeln("- Documents that need to be searchable and reusable.");
        builder.Writeln();

        // Explain which PDF/A version to choose.
        builder.Writeln("Choosing a PDF/A version:");
        builder.Writeln("PDF/A-1b: Preserve visual appearance only. Suitable for simple archival.");
        builder.Writeln("PDF/A-1a: Preserve appearance + document structure (tagged). Good for searchable archives.");
        builder.Writeln("PDF/A-2u: Unicode text extraction + visual preservation. Use when you need reliable text extraction.");
        builder.Writeln("PDF/A-3u: Same as PDF/A-2u but allows embedding attachments.");
        builder.Writeln("PDF/A-4: Latest standard, combines benefits of earlier versions with improved accessibility.");
        builder.Writeln();

        // Example: save the document as PDF/A-2u.
        builder.Writeln("Example: Saving this document as PDF/A-2u.");

        // Configure PDF save options to comply with PDF/A-2u.
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.Compliance = PdfCompliance.PdfA2u;

        // Save the document as a PDF file with the specified compliance.
        doc.Save("PdfA2uExample.pdf", saveOptions);
    }
}
