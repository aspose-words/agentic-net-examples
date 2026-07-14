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

        // Create sample documents with tables if they do not already exist.
        CreateSampleDocument(Path.Combine(inputDir, "Sample1.docx"));
        CreateSampleDocument(Path.Combine(inputDir, "Sample2.docx"));

        // Process each .docx file in the input directory.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Ensure the predefined table style exists in the document.
            const string styleName = "MyPredefinedTableStyle";
            TableStyle tableStyle;
            if (!doc.Styles.Any(s => s.Name == styleName))
            {
                // Add a new table style.
                tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, styleName);
                // Configure style properties (example settings).
                tableStyle.Borders.Color = System.Drawing.Color.Blue;
                tableStyle.Borders.LineStyle = LineStyle.Single;
                tableStyle.Borders.LineWidth = 1.0;
                tableStyle.Shading.BackgroundPatternColor = System.Drawing.Color.LightYellow;
                tableStyle.CellSpacing = 5.0;
                tableStyle.BottomPadding = 5.0;
                tableStyle.TopPadding = 5.0;
                tableStyle.LeftPadding = 5.0;
                tableStyle.RightPadding = 5.0;
            }
            else
            {
                tableStyle = (TableStyle)doc.Styles[styleName];
            }

            // Iterate through all tables in the document.
            NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
            foreach (Table table in tables)
            {
                // Apply the predefined style.
                table.Style = tableStyle;

                // Apply margin settings (example: left indent of 20 points).
                table.LeftIndent = 20.0;

                // Optionally adjust other table formatting.
                table.CellSpacing = 5.0;
                table.BottomPadding = 5.0;
                table.TopPadding = 5.0;
                table.LeftPadding = 5.0;
                table.RightPadding = 5.0;
            }

            // Save the modified document to the output directory.
            string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));
            doc.Save(outputPath);
        }
    }

    // Helper method to create a simple document containing a single table.
    private static void CreateSampleDocument(string filePath)
    {
        // If the file already exists, skip creation.
        if (File.Exists(filePath))
            return;

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table with two rows and two columns.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndRow();

        builder.EndTable();

        // Save the sample document.
        doc.Save(filePath);
    }
}
