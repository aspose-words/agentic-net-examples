using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Define folders for input and output documents.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputDocs");

        // Ensure the directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a few sample documents that contain tables.
        CreateSampleDocument(Path.Combine(inputDir, "Sample1.docx"));
        CreateSampleDocument(Path.Combine(inputDir, "Sample2.docx"));

        // Define a predefined table style that will be applied to every table.
        // The style will have a simple border, shading and cell spacing.
        TableStyle predefinedStyle = null;

        // Process each document in the input folder.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Ensure the predefined style exists in the current document.
            // If it does not exist, create it.
            if (predefinedStyle == null || predefinedStyle.Document != doc)
            {
                predefinedStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyPredefinedTableStyle");
                predefinedStyle.Borders.Color = System.Drawing.Color.Blue;
                predefinedStyle.Borders.LineStyle = LineStyle.Single;
                predefinedStyle.Borders.LineWidth = 1.0;
                predefinedStyle.Shading.BackgroundPatternColor = System.Drawing.Color.LightYellow;
                predefinedStyle.CellSpacing = 5.0;
                predefinedStyle.LeftPadding = 10.0;
                predefinedStyle.RightPadding = 10.0;
                predefinedStyle.TopPadding = 8.0;
                predefinedStyle.BottomPadding = 8.0;
            }

            // Iterate over all tables in the document.
            NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
            foreach (Table table in tables)
            {
                // Apply the predefined style.
                table.Style = predefinedStyle;

                // Set additional margin/indent settings for the table.
                table.LeftIndent = 20.0;          // Indent from the left page margin.
                table.CellSpacing = 5.0;          // Space between cells.
                table.TopPadding = 10.0;          // Padding above cell contents.
                table.BottomPadding = 10.0;       // Padding below cell contents.
                table.LeftPadding = 10.0;         // Padding to the left of cell contents.
                table.RightPadding = 10.0;        // Padding to the right of cell contents.
            }

            // Save the modified document to the output folder, preserving the original file name.
            string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));
            doc.Save(outputPath);
        }

        // Inform the user that processing is complete.
        Console.WriteLine($"Processed {Directory.GetFiles(inputDir, "*.docx").Length} document(s).");
        Console.WriteLine($"Modified files are saved in: {outputDir}");
    }

    // Helper method to create a simple document containing a single table.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table with three rows and two columns.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        // Third row.
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");
        builder.EndRow();

        builder.EndTable();

        // Save the sample document.
        doc.Save(filePath);
    }
}
