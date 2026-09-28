using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a temporary folder for sample documents.
        string folderPath = Path.Combine(Path.GetTempPath(), "AsposeBatchTables");
        Directory.CreateDirectory(folderPath);

        // Generate sample documents containing tables.
        for (int i = 1; i <= 3; i++)
        {
            string samplePath = Path.Combine(folderPath, $"Doc{i}.docx");
            CreateSampleDocument(samplePath);
        }

        // Process each document: replace all tables with a predefined style and margin settings.
        foreach (string filePath in Directory.GetFiles(folderPath, "*.docx"))
        {
            Document doc = new Document(filePath);

            // Ensure the predefined table style exists.
            const string styleName = "MyTableStyle";
            EnsureTableStyle(doc, styleName);

            // Apply the style and left margin to every table in the document.
            // Right margin is not directly supported; left indent is applied uniformly.
            const double marginPoints = 5.0 * 2.83464566929134; // 5 mm → points
            foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
            {
                table.Style = doc.Styles[styleName];
                table.LeftIndent = marginPoints;
            }

            // Save the processed document with a new name.
            string processedPath = Path.Combine(folderPath, $"Processed_{Path.GetFileName(filePath)}");
            doc.Save(processedPath);
        }

        // No interactive output required.
    }

    // Creates a simple document with a single 3x3 table.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.StartTable();
        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Writeln($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        doc.Save(filePath);
    }

    // Ensures that a table style with the given name exists in the document.
    private static void EnsureTableStyle(Document doc, string styleName)
    {
        // Return if the style already exists.
        if (doc.Styles[styleName] != null && doc.Styles[styleName].Type == StyleType.Table)
            return;

        // Create a new table style.
        TableStyle style = (TableStyle)doc.Styles.Add(StyleType.Table, styleName);
        style.Font.Name = "Arial";
        style.Font.Size = 10;
        style.Shading.BackgroundPatternColor = Color.LightGray;
        style.Borders.Color = Color.Black;
        style.Borders.LineWidth = 0.5;
    }
}
