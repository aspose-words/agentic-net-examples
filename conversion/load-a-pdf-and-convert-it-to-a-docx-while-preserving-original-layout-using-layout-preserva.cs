using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample PDF document.
        Document sourcePdf = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourcePdf);
        builder.Writeln("Sample PDF content for conversion.");
        sourcePdf.Save("sample.pdf", SaveFormat.Pdf);

        // Step 2: Load the PDF that was just created.
        Document pdfDocument = new Document("sample.pdf");

        // Optional: Ensure the layout is up‑to‑date before conversion.
        pdfDocument.UpdatePageLayout();

        // Step 3: Convert the PDF to DOCX while preserving layout.
        // No special save options are required for layout preservation in this scenario.
        pdfDocument.Save("converted.docx", SaveFormat.Docx);

        // Step 4: Validate that the DOCX file was created.
        if (!File.Exists("converted.docx"))
            throw new InvalidOperationException("The DOCX output file was not created.");

        // The program finishes automatically.
    }
}
