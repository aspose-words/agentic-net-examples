using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names for the sample PDF and the resulting XLSX.
        const string pdfPath = "sample_input.pdf";
        const string xlsxPath = "extracted_tables.xlsx";

        // -----------------------------------------------------------------
        // Step 1: Create a sample PDF containing a table.
        // -----------------------------------------------------------------
        Document pdfSource = new Document();
        DocumentBuilder builder = new DocumentBuilder(pdfSource);

        // Insert a simple 2x3 table with deterministic content.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Row 1, Col 1");
        builder.InsertCell();
        builder.Write("Row 1, Col 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Row 2, Col 1");
        builder.InsertCell();
        builder.Write("Row 2, Col 2");
        builder.EndRow();

        builder.EndTable();

        // Save the document as PDF.
        pdfSource.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException($"Failed to create the PDF file at '{pdfPath}'.");

        // -----------------------------------------------------------------
        // Step 2: Load the PDF and convert it to XLSX.
        // -----------------------------------------------------------------
        Document pdfDocument = new Document(pdfPath);

        // Configure XLSX save options to place each section on a separate worksheet.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            SectionMode = XlsxSectionMode.MultipleWorksheets,
            SaveFormat = SaveFormat.Xlsx
        };

        // Save the loaded PDF as an XLSX spreadsheet.
        pdfDocument.Save(xlsxPath, xlsxOptions);

        // Verify that the XLSX file was created.
        if (!File.Exists(xlsxPath))
            throw new InvalidOperationException($"Failed to create the XLSX file at '{xlsxPath}'.");

        // Indicate successful completion.
        Console.WriteLine("PDF successfully converted to XLSX with tables extracted to worksheets.");
    }
}
