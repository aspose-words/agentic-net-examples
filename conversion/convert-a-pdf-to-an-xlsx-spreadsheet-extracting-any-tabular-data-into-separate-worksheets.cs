using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string pdfPath = "input.pdf";
        const string xlsxPath = "output.xlsx";

        // -----------------------------------------------------------------
        // 1. Create a sample Word document containing a table.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Insert a simple 2x3 table with sample data.
        Table table = builder.StartTable();

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

        // Save the document as PDF – this will be our input file.
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The input PDF was not created.");

        // -----------------------------------------------------------------
        // 2. Load the PDF and convert it to XLSX, placing each section on a separate worksheet.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            // Each section of the document becomes a separate worksheet.
            SectionMode = XlsxSectionMode.MultipleWorksheets,
            SaveFormat = SaveFormat.Xlsx
        };

        pdfDoc.Save(xlsxPath, xlsxOptions);

        // -----------------------------------------------------------------
        // 3. Validate that the XLSX file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(xlsxPath))
            throw new InvalidOperationException("The output XLSX was not created.");

        Console.WriteLine("PDF successfully converted to XLSX: " + Path.GetFullPath(xlsxPath));
    }
}
