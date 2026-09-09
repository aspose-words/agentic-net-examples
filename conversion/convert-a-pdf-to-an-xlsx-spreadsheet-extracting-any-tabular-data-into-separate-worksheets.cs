using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary files.
        const string pdfPath = "input.pdf";
        const string xlsxPath = "output.xlsx";

        // -------------------------------------------------
        // 1. Create a sample document containing two tables,
        //    each placed in its own section.
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First section with the first table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("A1");
        builder.InsertCell();
        builder.Write("B1");
        builder.EndRow();
        builder.EndTable();

        // Insert a section break so the next table is in a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second section with the second table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Col 1");
        builder.InsertCell();
        builder.Write("Col 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("C1");
        builder.InsertCell();
        builder.Write("D1");
        builder.EndRow();
        builder.EndTable();

        // Save the document as PDF – this will be the input for conversion.
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // -------------------------------------------------
        // 2. Load the PDF and convert it to XLSX.
        // -------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // Configure XLSX save options to create a separate worksheet per section.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            SectionMode = XlsxSectionMode.MultipleWorksheets,
            SaveFormat = SaveFormat.Xlsx
        };

        pdfDoc.Save(xlsxPath, xlsxOptions);

        // -------------------------------------------------
        // 3. Validate that the XLSX file was created.
        // -------------------------------------------------
        if (!File.Exists(xlsxPath))
            throw new InvalidOperationException($"The expected output file '{xlsxPath}' was not created.");

        Console.WriteLine($"PDF successfully converted to XLSX. Output file: {Path.GetFullPath(xlsxPath)}");
    }
}
