using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing; // For LineStyle enum
using System.Drawing; // For Color

class Program
{
    static void Main()
    {
        // Use folders relative to the executable location.
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string inputFolder = Path.Combine(baseDir, "Input");
        string outputFolder = Path.Combine(baseDir, "Output");

        // Ensure the directories exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // If there are no .docx files, create a sample document with a table.
        if (Directory.GetFiles(inputFolder, "*.docx").Length == 0)
        {
            var sampleDoc = new Document();
            var builder = new DocumentBuilder(sampleDoc);
            builder.Writeln("Sample document created by the batch processor.");
            builder.StartTable();
            for (int r = 0; r < 3; r++)
            {
                for (int c = 0; c < 4; c++)
                {
                    builder.InsertCell();
                    builder.Writeln($"R{r + 1}C{c + 1}");
                }
                builder.EndRow();
            }
            builder.EndTable();
            string samplePath = Path.Combine(inputFolder, "Sample.docx");
            sampleDoc.Save(samplePath);
        }

        const string predefinedStyleName = "MyBatchTableStyle";

        foreach (string filePath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(filePath);

            // Create (or retrieve) the predefined table style.
            TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, predefinedStyleName);
            tableStyle.Borders.Color = Color.Blue;
            tableStyle.Borders.LineStyle = LineStyle.Single;
            tableStyle.Borders.LineWidth = 1.0; // points
            tableStyle.CellSpacing = 0;
            tableStyle.LeftPadding = 5;
            tableStyle.RightPadding = 5;
            tableStyle.TopPadding = 5;
            tableStyle.BottomPadding = 5;
            tableStyle.Shading.BackgroundPatternColor = Color.LightGray;

            // Apply the style to every table.
            NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
            foreach (Table table in tables)
            {
                table.Style = tableStyle;
                table.DistanceTop = 12;    // points
                table.DistanceBottom = 12; // points
                table.DistanceLeft = 12;   // points
                table.DistanceRight = 12;  // points
                table.LeftIndent = 0;
            }

            // Convert style formatting to direct formatting.
            doc.ExpandTableStylesToDirectFormatting();

            // Save the modified document.
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
            doc.Save(outputPath);
        }

        Console.WriteLine("Processing complete. Check the 'Output' folder.");
    }
}
