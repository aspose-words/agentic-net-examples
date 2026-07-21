using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF file.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample PDF content for conversion.");
        source.Save("sample.pdf", SaveFormat.Pdf);

        // Load the PDF with load options (default options preserve layout).
        PdfLoadOptions loadOptions = new PdfLoadOptions();
        Document pdfDoc = new Document("sample.pdf", loadOptions);

        // Convert the PDF to DOCX while preserving the original layout.
        pdfDoc.Save("converted.docx", SaveFormat.Docx);

        // Verify that the DOCX file was created.
        if (!File.Exists("converted.docx"))
            throw new InvalidOperationException("The DOCX output file was not created.");
    }
}
