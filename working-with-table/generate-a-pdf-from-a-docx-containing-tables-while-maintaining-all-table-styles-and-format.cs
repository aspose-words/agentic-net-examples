using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to construct a table with a built‑in style.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table.
        Table table = builder.StartTable();

        // First row – header cells.
        builder.InsertCell();
        builder.Write("Product");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // Second row – data.
        builder.InsertCell();
        builder.Write("Apples");
        builder.InsertCell();
        builder.Write("10");
        builder.EndRow();

        // Third row – data.
        builder.InsertCell();
        builder.Write("Bananas");
        builder.InsertCell();
        builder.Write("20");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply a built‑in table style that includes shading and borders.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Convert any style‑based formatting to direct formatting.
        // This ensures the PDF retains the exact appearance of the table.
        doc.ExpandTableStylesToDirectFormatting();

        // Save the document as DOCX (optional, for verification).
        string docxPath = Path.Combine(outputDir, "SampleTable.docx");
        doc.Save(docxPath);

        // Save the document as PDF, preserving table styles and formatting.
        string pdfPath = Path.Combine(outputDir, "SampleTable.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Simple validation – ensure the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // Inform the user (no interactive prompts required).
        Console.WriteLine($"PDF generated successfully at: {pdfPath}");
    }
}
