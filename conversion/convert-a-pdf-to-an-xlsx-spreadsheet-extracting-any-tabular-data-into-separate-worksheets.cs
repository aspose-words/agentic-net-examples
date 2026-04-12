using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class PdfToXlsxConverter
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string pdfPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.pdf");
        string xlsxPath = Path.Combine(Directory.GetCurrentDirectory(), "converted.xlsx");

        // -----------------------------------------------------------------
        // 1. Create a sample PDF containing a table.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple 3x3 table.
        builder.StartTable();
        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Writeln($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Save the document as PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 2. Load the PDF and convert it to XLSX.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // Configure XLSX save options to create a separate worksheet for each section.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            SectionMode = XlsxSectionMode.MultipleWorksheets,
            SaveFormat = SaveFormat.Xlsx
        };

        // Save the loaded PDF as an XLSX spreadsheet.
        pdfDoc.Save(xlsxPath, xlsxOptions);

        // -----------------------------------------------------------------
        // 3. Validate that the XLSX file was created successfully.
        // -----------------------------------------------------------------
        if (!File.Exists(xlsxPath))
        {
            throw new FileNotFoundException("The XLSX output file was not created.", xlsxPath);
        }

        FileInfo info = new FileInfo(xlsxPath);
        if (info.Length == 0)
        {
            throw new InvalidOperationException("The XLSX output file is empty.");
        }

        // Indicate successful completion (optional logging).
        Console.WriteLine($"PDF successfully converted to XLSX. Output file: {xlsxPath}");
    }
}
