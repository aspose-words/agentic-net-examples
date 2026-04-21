using System;
using System.Collections.Generic;
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

        // Fixed number of columns for the table.
        const int columnCount = 3;

        // Start building the table.
        Table table = builder.StartTable();

        // Header row.
        string[] headers = { "ID", "Name", "Score" };
        foreach (string header in headers)
        {
            builder.InsertCell();
            builder.Write(header);
        }
        builder.EndRow();

        // Sample data source collection.
        var data = new List<(int Id, string Name, double Score)>
        {
            (1, "Alice", 85.5),
            (2, "Bob", 92.0),
            (3, "Charlie", 78.3)
        };

        // Dynamically add a row for each item in the collection.
        foreach (var item in data)
        {
            builder.InsertCell();
            builder.Write(item.Id.ToString());

            builder.InsertCell();
            builder.Write(item.Name);

            builder.InsertCell();
            builder.Write(item.Score.ToString("F1"));

            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Adjust column widths to fit the content.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OutputTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file at {outputPath}");
        }
    }
}
