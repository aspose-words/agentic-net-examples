using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare folders for input and output documents.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a couple of sample documents that contain tables.
        CreateSampleDocument(Path.Combine(inputDir, "Sample1.docx"));
        CreateSampleDocument(Path.Combine(inputDir, "Sample2.docx"));

        // Process each document: replace all tables with a predefined style and margin settings.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Define a custom table style that will be applied to every table.
            TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyTableStyle");
            tableStyle.Shading.BackgroundPatternColor = Color.LightGray;
            tableStyle.Borders.Color = Color.Black;
            tableStyle.Borders.LineStyle = LineStyle.Single;
            tableStyle.Borders.LineWidth = 1.0;
            tableStyle.CellSpacing = 5;          // Space between cells.
            tableStyle.LeftPadding = 10;         // Padding inside cells.
            tableStyle.RightPadding = 10;
            tableStyle.TopPadding = 5;
            tableStyle.BottomPadding = 5;

            // Apply the style and margin settings to every table in the document.
            NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
            foreach (Table tbl in tables)
            {
                tbl.Style = tableStyle;          // Apply the predefined style.
                tbl.LeftIndent = 20;             // Example left indent (acts like a margin for the table).
            }

            // Save the processed document to the output folder.
            string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));
            doc.Save(outputPath);
        }
    }

    // Helper method to create a simple document containing a table.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a 2x2 table with placeholder text.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();
        builder.EndTable();

        doc.Save(filePath);
    }
}
