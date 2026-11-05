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

        // Insert a table with 3 columns and 5 rows.
        Table table = builder.StartTable();

        // Define header row.
        for (int col = 0; col < 3; col++)
        {
            builder.InsertCell();
            builder.Write($"Header {col + 1}");
        }
        builder.EndRow();

        // Fill table with sample data.
        for (int row = 0; row < 4; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                // Vary content length to demonstrate width adjustment.
                string content = $"Row {row + 1}, Column {col + 1}";
                if (col == 1) content += " - longer text example";
                builder.Write(content);
            }
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Adjust column widths proportionally to fit the content.
        // AutoFitToContents resizes each column based on its longest cell content.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document to a file.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "AdjustedTable.docx");
        doc.Save(outputPath);
    }
}
