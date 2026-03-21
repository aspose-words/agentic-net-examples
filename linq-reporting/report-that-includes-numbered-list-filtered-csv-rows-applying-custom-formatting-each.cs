using System;
using System.Data;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

class CsvReportGenerator
{
    static void Main()
    {
        // Create a temporary CSV file with sample data.
        string csvPath = Path.Combine(Path.GetTempPath(), "people.csv");
        if (!File.Exists(csvPath))
        {
            File.WriteAllLines(csvPath, new[]
            {
                "Name,Age,City",
                "Alice,28,New York",
                "Bob,35,London",
                "Charlie,32,Sydney",
                "Diana,24,Tokyo"
            });
        }

        // Load CSV data into a DataTable.
        DataTable table = LoadCsvIntoDataTable(csvPath, hasHeaders: true);

        // Ensure the Age column exists and convert it to int for proper filtering.
        if (table.Columns.Contains("Age"))
        {
            // Change the column type to int.
            table.Columns["Age"].DataType = typeof(int);

            foreach (DataRow row in table.Rows)
            {
                if (int.TryParse(row["Age"]?.ToString(), out int age))
                    row["Age"] = age;
                else
                    row["Age"] = DBNull.Value;
            }
        }
        else
        {
            Console.WriteLine("Column 'Age' not found in the CSV data.");
            return;
        }

        // Filter rows where Age >= 30 using LINQ.
        var filteredRows = table.AsEnumerable()
                                .Where(r => r.Field<int?>("Age") >= 30)
                                .ToList();

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a numbered list.
        builder.ListFormat.ApplyNumberDefault();

        // Iterate over the filtered rows and add each as a list item.
        foreach (DataRow row in filteredRows)
        {
            // Apply custom formatting: make the Name bold, Age italic.
            builder.Font.Bold = true;
            builder.Write(row["Name"].ToString());

            builder.Font.Bold = false;
            builder.Write(" - ");

            builder.Font.Italic = true;
            builder.Writeln(row["Age"].ToString());
            builder.Font.Italic = false;
        }

        // End the numbered list.
        builder.ListFormat.RemoveNumbers();

        // Save the resulting document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FilteredReport.docx");
        doc.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
    }

    // Helper method to read a CSV file into a DataTable.
    private static DataTable LoadCsvIntoDataTable(string path, bool hasHeaders)
    {
        DataTable dt = new DataTable();

        using (var reader = new StreamReader(path))
        {
            bool firstLine = true;
            while (!reader.EndOfStream)
            {
                string line = reader.ReadLine();
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                string[] fields = line.Split(',');

                if (firstLine && hasHeaders)
                {
                    foreach (string header in fields)
                        dt.Columns.Add(header.Trim());
                    firstLine = false;
                    continue;
                }

                if (dt.Columns.Count == 0)
                {
                    for (int i = 0; i < fields.Length; i++)
                        dt.Columns.Add($"Column{i + 1}");
                }

                DataRow dr = dt.NewRow();
                for (int i = 0; i < fields.Length; i++)
                    dr[i] = fields[i].Trim();
                dt.Rows.Add(dr);
                firstLine = false;
            }
        }

        return dt;
    }
}
