using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Sample data source: each string array represents a row.
        var data = new List<string[]>
        {
            new[] { "Alice", "Engineering", "1000" },
            new[] { "Bob", "Marketing", "1500" },
            new[] { "Charlie", "HR", "1200" }
        };

        // Create a new blank document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Fixed number of columns for the table.
        const int columnCount = 3;

        // Start the table.
        Table table = builder.StartTable();

        // Add a header row.
        string[] headers = { "Name", "Department", "Salary" };
        for (int i = 0; i < columnCount; i++)
        {
            builder.InsertCell();
            builder.Write(headers[i]);
        }
        builder.EndRow();

        // Add rows from the data source.
        foreach (var row in data)
        {
            for (int i = 0; i < columnCount; i++)
            {
                builder.InsertCell();
                builder.Write(row[i]);
            }
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DynamicTable.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"The output file was not created: {outputPath}");
        }
    }
}
