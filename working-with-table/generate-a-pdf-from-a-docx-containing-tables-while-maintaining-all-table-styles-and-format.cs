using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Define output folder and ensure it exists.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Paths for the intermediate DOCX and final PDF.
        string docPath = Path.Combine(artifactsDir, "Sample.docx");
        string pdfPath = Path.Combine(artifactsDir, "Sample.pdf");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and keep a reference to it.
        Table table = builder.StartTable();

        // First row – header cells.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Second row – data cells.
        builder.InsertCell();
        builder.Write("Data 1");
        builder.InsertCell();
        builder.Write("Data 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Create a custom table style and configure some formatting.
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomTableStyle");
        tableStyle.RowStripe = 1; // Apply row banding.
        tableStyle.CellSpacing = 5; // Space between cells.
        tableStyle.Shading.BackgroundPatternColor = Color.LightBlue;
        tableStyle.Borders.Color = Color.DarkBlue;
        tableStyle.Borders.LineStyle = LineStyle.Single;

        // Apply the style to the table.
        table.Style = tableStyle;

        // Convert style-based formatting to direct formatting so it is preserved in PDF.
        doc.ExpandTableStylesToDirectFormatting();

        // Save the document as DOCX (optional, shows the intermediate file).
        doc.Save(docPath);

        // Save the same document as PDF, preserving all table formatting.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new Exception("PDF file was not created.");

        // Optionally, inform that the process completed (no interactive input required).
        Console.WriteLine("PDF generated successfully at: " + pdfPath);
    }
}
