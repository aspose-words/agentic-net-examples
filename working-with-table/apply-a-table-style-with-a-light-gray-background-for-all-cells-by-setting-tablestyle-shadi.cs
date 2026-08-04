using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a 2x2 table.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndTable();

        // Create a custom table style.
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "LightGrayStyle");

        // Apply a light gray background to all cells via the style's shading.
        tableStyle.Shading.BackgroundPatternColor = Color.LightGray;

        // Assign the custom style to the table.
        table.Style = tableStyle;

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithLightGrayStyle.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Optionally, you could open the file automatically (commented out to avoid side effects).
        // System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outputPath) { UseShellExecute = true });
    }
}
