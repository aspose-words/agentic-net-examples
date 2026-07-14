using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF document.
        Document sourceDocument = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDocument);
        builder.Writeln("Sample PDF content for conversion.");
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndTable();

        // Save the sample as PDF.
        string pdfPath = "sample.pdf";
        sourceDocument.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Load the PDF (layout‑preservation is the default behavior).
        PdfLoadOptions loadOptions = new PdfLoadOptions();
        Document pdfDocument = new Document(pdfPath, loadOptions);

        // Convert the loaded PDF to DOCX while preserving layout.
        string docxPath = "converted.docx";
        pdfDocument.Save(docxPath, SaveFormat.Docx);

        // Verify that the DOCX file was created.
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("The DOCX file was not created.");
    }
}
