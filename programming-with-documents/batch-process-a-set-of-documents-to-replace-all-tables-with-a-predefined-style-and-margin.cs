using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a temporary working folder.
        string workFolder = Path.Combine(Path.GetTempPath(), "AsposeBatchDemo");
        Directory.CreateDirectory(workFolder);

        // Create sample documents with tables.
        for (int i = 1; i <= 3; i++)
        {
            string samplePath = Path.Combine(workFolder, $"Doc{i}.docx");
            CreateSampleDocumentWithTable(samplePath);
        }

        // Define the style and margin settings to apply.
        StyleIdentifier targetStyle = StyleIdentifier.TableGrid; // Predefined style.
        double leftMarginPoints = 20.0;   // Left margin in points.
        double rightMarginPoints = 20.0;  // Right margin in points.

        // Process each document: apply the predefined style and margins to all tables.
        foreach (string filePath in Directory.GetFiles(workFolder, "*.docx"))
        {
            Document doc = new Document(filePath);

            // Find all tables in the document.
            NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
            foreach (Table table in tables)
            {
                // Apply the predefined style.
                table.StyleIdentifier = targetStyle;

                // Apply left indent (margin).
                table.LeftIndent = leftMarginPoints;

                // Aspose.Words Table does not expose a RightIndent property.
                // Use RightPadding to simulate a right margin for the table.
                table.RightPadding = rightMarginPoints;
            }

            // Save the processed document (overwrite original).
            doc.Save(filePath);
        }

        // Optional verification that files exist after processing.
        bool allProcessed = Directory.GetFiles(workFolder, "*.docx")
            .All(f => File.Exists(f));
    }

    // Helper method to create a simple document containing a table.
    private static void CreateSampleDocumentWithTable(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a heading.
        builder.Writeln($"Sample Document: {Path.GetFileNameWithoutExtension(filePath)}");

        // Insert a 3x3 table.
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

        // Save the document.
        doc.Save(filePath);
    }
}
