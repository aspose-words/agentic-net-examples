using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    // Predefined style name that will be applied to every table.
    private const string TableStyleName = "MyPredefinedTableStyle";

    // Margin (left indent) that will be applied to every table, measured in points.
    private const double TableLeftIndent = 20.0;

    public static void Main()
    {
        // Directories for sample source documents and processed output.
        string sourceDir = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Processed");

        Directory.CreateDirectory(sourceDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Step 1: Create a few sample documents that contain tables.
        // -----------------------------------------------------------------
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);

            // Add a simple 2x2 table with some text.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write($"Doc {i} - Row 1, Cell 1");
            builder.InsertCell();
            builder.Write($"Doc {i} - Row 1, Cell 2");
            builder.EndRow();

            builder.InsertCell();
            builder.Write($"Doc {i} - Row 2, Cell 1");
            builder.InsertCell();
            builder.Write($"Doc {i} - Row 2, Cell 2");
            builder.EndRow();

            builder.EndTable();

            // Save the sample document.
            string samplePath = Path.Combine(sourceDir, $"Sample{i}.docx");
            sampleDoc.Save(samplePath);
        }

        // -----------------------------------------------------------------
        // Step 2: Process each document – replace all tables with the
        // predefined style and apply the left‑indent margin.
        // -----------------------------------------------------------------
        foreach (string filePath in Directory.GetFiles(sourceDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Ensure the predefined table style exists in the current document.
            EnsureTableStyleExists(doc);

            // Iterate over all tables in the document.
            NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
            foreach (Table tbl in tables)
            {
                // Apply the predefined style.
                tbl.StyleName = TableStyleName;

                // Apply the left‑indent margin.
                tbl.LeftIndent = TableLeftIndent;

                // Optional: expand style to direct formatting if further inspection is needed.
                // doc.ExpandTableStylesToDirectFormatting(); // Uncomment if required.
            }

            // Save the processed document to the output folder.
            string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));
            doc.Save(outputPath);
        }

        // The program finishes automatically; no user interaction is required.
    }

    /// <summary>
    /// Adds the predefined table style to the document if it does not already exist.
    /// </summary>
    /// <param name="doc">The document to which the style will be added.</param>
    private static void EnsureTableStyleExists(Document doc)
    {
        // Check whether the style already exists.
        if (doc.Styles[TableStyleName] != null)
            return;

        // Create a new table style.
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, TableStyleName);
        tableStyle.RowStripe = 2;                     // Apply row banding.
        tableStyle.CellSpacing = 5;                    // Space between cells (points).
        tableStyle.BottomPadding = 5;
        tableStyle.TopPadding = 5;
        tableStyle.LeftPadding = 5;
        tableStyle.RightPadding = 5;
        tableStyle.Shading.BackgroundPatternColor = Color.AntiqueWhite;
        tableStyle.Borders.Color = Color.Blue;
        tableStyle.Borders.LineStyle = LineStyle.DotDash;
        tableStyle.Borders.LineWidth = 0.5;
    }
}
