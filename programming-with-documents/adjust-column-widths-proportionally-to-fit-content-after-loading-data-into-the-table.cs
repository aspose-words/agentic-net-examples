using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table with 3 columns.
        Table table = builder.StartTable();

        // Add header cells.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertCell();
            builder.Write($"Header {i + 1}");
        }
        builder.EndRow();

        // Sample data to populate the table.
        string[,] data = new string[,]
        {
            { "Short", "Medium length", "A very long piece of text that should expand the column" },
            { "Another short", "Another medium length text", "Another long text that will affect column width" },
            { "Tiny", "Some text", "Longer text again" },
            { "Small", "Medium", "Very very long text that will definitely cause the column to expand" }
        };

        // Fill the table with data rows.
        for (int r = 0; r < data.GetLength(0); r++)
        {
            for (int c = 0; c < data.GetLength(1); c++)
            {
                builder.InsertCell();
                builder.Write(data[r, c]);
            }
            builder.EndRow();
        }

        // End the table construction.
        builder.EndTable();

        // Adjust column widths proportionally to fit the content.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document to disk.
        string outputPath = "AdjustedTable.docx";
        doc.Save(outputPath);
    }
}
