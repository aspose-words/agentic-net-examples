using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Use DocumentBuilder to add content describing the two PDF standards.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Title.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("PDF Standards Comparison");

        // Section for PDF 1.7.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("PDF 1.7 (ISO 32000-1)");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("PDF 1.7 is the original PDF specification defined by ISO 32000-1. " +
                         "It is widely supported and forms the basis for most PDF viewers and creators.");

        // Section for PDF 2.0.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("PDF 2.0 (ISO 32000-2)");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("PDF 2.0 is the newer specification defined by ISO 32000-2. " +
                         "It adds new features such as enhanced encryption, richer metadata, " +
                         "and improved handling of color spaces while maintaining backward compatibility.");

        // Rebuild the page layout before rendering to PDF.
        doc.UpdatePageLayout();

        // Save the document as PDF 1.7 compliant.
        PdfSaveOptions saveOptions17 = new PdfSaveOptions();
        saveOptions17.Compliance = PdfCompliance.Pdf17;
        doc.Save("PdfComparison_Pdf17.pdf", saveOptions17);

        // Save the same document as PDF 2.0 compliant.
        PdfSaveOptions saveOptions20 = new PdfSaveOptions();
        saveOptions20.Compliance = PdfCompliance.Pdf20;
        doc.Save("PdfComparison_Pdf20.pdf", saveOptions20);
    }
}
