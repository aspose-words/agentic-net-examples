using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the intermediate DOCX and final HTML files.
        string docPath = Path.Combine(outputDir, "ComplexTable.docx");
        string htmlPath = Path.Combine(outputDir, "ComplexTable.html");

        // Create a new document and a builder to construct a complex table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table.
        Table table = builder.StartTable();

        // Ensure the table has at least one row before applying formatting.
        table.EnsureMinimum();

        // Apply borders to the table (outline and inside borders).
        table.SetBorder(BorderType.Left, LineStyle.Single, 1.5, Color.Black, true);
        table.SetBorder(BorderType.Right, LineStyle.Single, 1.5, Color.Black, true);
        table.SetBorder(BorderType.Top, LineStyle.Single, 1.5, Color.Black, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Single, 1.5, Color.Black, true);
        table.SetBorder(BorderType.Horizontal, LineStyle.Single, 1.0, Color.Gray, true);
        table.SetBorder(BorderType.Vertical, LineStyle.Single, 1.0, Color.Gray, true);

        // Set cell padding for the whole table.
        table.LeftPadding = 10;
        table.RightPadding = 10;
        table.TopPadding = 5;
        table.BottomPadding = 5;

        // First row – header cells.
        builder.InsertCell();
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Font.Bold = true;
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Font.Bold = false;
        builder.Write("Header 2");
        builder.EndRow();

        // Second row – normal data cells.
        builder.InsertCell();
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
        builder.Write("Data 1");
        builder.InsertCell();
        builder.Write("Data 2");
        builder.EndRow();

        // Third row – demonstrate a merged cell (horizontal merge).
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Merged Cell");
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous; // Continue merge.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the intermediate DOCX (optional, shows the source document).
        doc.Save(docPath);

        // Convert the document to HTML while preserving borders and padding.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            TableWidthOutputMode = HtmlElementSizeOutputMode.All
        };
        doc.Save(htmlPath, htmlOptions);

        // Validate that the HTML file was created.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML conversion failed: output file not found.");

        Console.WriteLine("Conversion completed successfully. HTML saved to:");
        Console.WriteLine(htmlPath);
    }
}
