using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

class CsvReportGenerator
{
    static void Main()
    {
        // Path to the source CSV file.
        string csvPath = @"C:\Data\people.csv";

        // Load CSV data into a DataTable.
        DataTable table = LoadCsvIntoDataTable(csvPath, hasHeaders: true);

        // Apply a filter to the DataTable (example: only rows where Age >= 30).
        DataView view = new DataView(table);
        view.RowFilter = "Age >= 30";

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a numbered list.
        builder.ListFormat.ApplyNumberDefault();

        // Iterate over the filtered rows and add each as a list item.
        foreach (DataRow row in view.ToTable().Rows)
        {
            // Apply custom formatting: make the Name bold, Age normal.
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

        // Save the resulting document.
        string outputPath = @"C:\Output\FilteredReport.docx";
        doc.Save(outputPath);
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
                    // Create columns from the header row.
                    foreach (string header in fields)
                        dt.Columns.Add(header.Trim());
                    firstLine = false;
                    continue;
                }

                // If the table has no columns yet (no header row), create generic column names.
                if (dt.Columns.Count == 0)
                {
                    for (int i = 0; i < fields.Length; i++)
                        dt.Columns.Add($"Column{i + 1}");
                }

                // Add the data row.
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
