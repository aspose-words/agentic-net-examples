using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF document.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample PDF content.");
        source.Save("sample.pdf", SaveFormat.Pdf);

        // Load the PDF document that was just created.
        Document pdfDoc = new Document("sample.pdf");

        // Ensure the layout is up‑to‑date before conversion.
        pdfDoc.UpdatePageLayout();

        // Convert the PDF to DOCX while preserving the original layout.
        pdfDoc.Save("output.docx", SaveFormat.Docx);

        // Verify that the DOCX file was created.
        if (!File.Exists("output.docx"))
            throw new InvalidOperationException("The DOCX output file was not created.");
    }
}
