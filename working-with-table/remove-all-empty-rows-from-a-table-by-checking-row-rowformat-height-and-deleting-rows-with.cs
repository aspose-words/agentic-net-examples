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

        // Start building a table.
        Table table = builder.StartTable();

        // Add five rows. Even-indexed rows will have zero height (considered empty).
        for (int i = 0; i < 5; i++)
        {
            // Set the height for the current row.
            // HeightRule.Exactly forces the row to use the specified height.
            builder.RowFormat.HeightRule = HeightRule.Exactly;
            builder.RowFormat.Height = (i % 2 == 0) ? 0 : 50; // Zero height for even rows.

            // Insert two cells with some text.
            builder.InsertCell();
            builder.Write($"Row {i} Cell 1");
            builder.InsertCell();
            builder.Write($"Row {i} Cell 2");

            // End the current row.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Remove rows that have zero height.
        // Iterate backwards to avoid index shifting when removing rows.
        for (int i = table.Rows.Count - 1; i >= 0; i--)
        {
            Row row = table.Rows[i];
            if (row.RowFormat.Height == 0)
            {
                row.Remove();
            }
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Result.docx");
        doc.Save(outputPath);

        // Simple verification that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
