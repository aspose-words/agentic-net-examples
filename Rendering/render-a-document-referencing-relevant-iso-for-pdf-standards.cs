using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfComplianceExample
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Use DocumentBuilder to add content that references relevant ISO standards.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This PDF document complies with the following ISO standards:");
        builder.Writeln("• ISO 32000-1: PDF 1.7");
        builder.Writeln("• ISO 32000-2: PDF 2.0");
        builder.Writeln("• ISO 19005-1: PDF/A-1 (PDF/A-1a, PDF/A-1b)");
        builder.Writeln("• ISO 19005-2: PDF/A-2 (PDF/A-2a, PDF/A-2u)");
        builder.Writeln("• ISO 19005-3: PDF/A-3 (PDF/A-3a, PDF/A-3u)");
        builder.Writeln("• ISO 19005-4: PDF/A-4 (PDF/A-4, PDF/A-4f)");
        builder.Writeln("• ISO 14289-1: PDF/UA-1");
        builder.Writeln("• ISO 14289-2: PDF/UA-2");

        // Create PdfSaveOptions to specify the desired PDF compliance level.
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Set the compliance to PDF/A-4 with PDF/UA-2 (ISO 19005-4 and ISO 14289-2).
        saveOptions.Compliance = PdfCompliance.PdfA4Ua2;

        // Export the document structure (tags) which is required for PDF/UA compliance.
        saveOptions.ExportDocumentStructure = true;

        // Save the document as a PDF file using the specified options.
        doc.Save("Output_Compliant.pdf", saveOptions);
    }
}
