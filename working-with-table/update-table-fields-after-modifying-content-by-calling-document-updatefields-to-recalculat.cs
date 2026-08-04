using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with a formula field that sums the values above it.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Price");
        builder.EndRow();

        // First data row.
        builder.InsertCell();
        builder.Write("Apple");
        builder.InsertCell();
        builder.Write("2");
        builder.EndRow();

        // Second data row.
        builder.InsertCell();
        builder.Write("Banana");
        builder.InsertCell();
        builder.Write("3");
        builder.EndRow();

        // Total row with a formula field that sums the column above.
        builder.InsertCell();
        builder.Write("Total");
        builder.InsertCell();
        // Insert a field that calculates the sum of the numeric cells above it.
        // Use the string overload of InsertField which updates the field automatically.
        builder.InsertField("=SUM(ABOVE)");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Modify the price of "Apple" from 2 to 5.
        // Locate the cell at row index 1 (second row), column index 1 (second column).
        Cell priceCell = table.Rows[1].Cells[1];
        // Clear existing content.
        priceCell.RemoveAllChildren();
        // Add a new paragraph with the updated price.
        priceCell.AppendChild(new Paragraph(doc));
        priceCell.FirstParagraph.AppendChild(new Run(doc, "5"));

        // Recalculate all fields in the document (the SUM field will now reflect the new total).
        doc.UpdateFields();

        // Define output path.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outputPath = Path.Combine(artifactsDir, "TableWithUpdatedFields.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file at '{outputPath}'.");
        }
    }
}
