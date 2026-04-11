using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string pdfPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.pdf");
        string docxPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.docx");
        string xlsxPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.xlsx");

        // -----------------------------------------------------------------
        // Step 1: Create a sample PDF containing a simple table.
        // -----------------------------------------------------------------
        Document pdfDocument = new Document();
        DocumentBuilder builder = new DocumentBuilder(pdfDocument);

        builder.Writeln("Sample PDF with a table:");
        // Build a 2x2 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndTable();

        // Save the document as PDF.
        pdfDocument.Save(pdfPath, SaveFormat.Pdf);

        // Verify PDF creation.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the PDF file.");

        // -----------------------------------------------------------------
        // Step 2: Load the PDF and convert it to DOCX.
        // -----------------------------------------------------------------
        Document loadedPdf = new Document(pdfPath);
        loadedPdf.Save(docxPath, SaveFormat.Docx);

        // Verify DOCX creation.
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("Failed to convert PDF to DOCX.");

        // -----------------------------------------------------------------
        // Step 3: Load the DOCX and convert it to XLSX (tables become worksheets).
        // -----------------------------------------------------------------
        Document loadedDocx = new Document(docxPath);
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            // Save all sections into a single worksheet (optional).
            SectionMode = XlsxSectionMode.SingleWorksheet,
            SaveFormat = SaveFormat.Xlsx
        };
        loadedDocx.Save(xlsxPath, xlsxOptions);

        // Verify XLSX creation.
        if (!File.Exists(xlsxPath))
            throw new InvalidOperationException("Failed to convert DOCX to XLSX.");

        // Indicate successful completion.
        Console.WriteLine("Conversion pipeline completed successfully.");
    }
}
