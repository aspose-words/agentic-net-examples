using System;
using System.Data;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

class CsvFilterExample
{
    static void Main()
    {
        // Load the Word template that contains the reporting tags.
        // Example template tag: <<foreach [persons]>><<[Name]>> <<[/foreach]>>
        Document doc = new Document("Template.docx");

        // --------------------------------------------------------------------
        // Load CSV data into a DataTable.
        // --------------------------------------------------------------------
        // The CSV file is expected to have a header row.
        // Example CSV content:
        // Name,Age,Status
        // John,30,Active
        // Jane,25,Inactive
        // --------------------------------------------------------------
        string csvPath = "People.csv";
        DataTable csvTable = new DataTable("persons");

        using (var reader = new StreamReader(csvPath))
        {
            bool isFirstLine = true;
            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine();
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                var fields = line.Split(',');

                if (isFirstLine)
                {
                    // Create columns from the header row.
                    foreach (var header in fields)
                        csvTable.Columns.Add(header.Trim());
                    isFirstLine = false;
                }
                else
                {
                    // Add data rows.
                    var row = csvTable.NewRow();
                    for (int i = 0; i < fields.Length; i++)
                        row[i] = fields[i].Trim();
                    csvTable.Rows.Add(row);
                }
            }
        }

        // --------------------------------------------------------------------
        // Filter rows where the "Status" column equals "Active" using LINQ Where.
        // --------------------------------------------------------------------
        var activeRows = csvTable.AsEnumerable()
                                 .Where(r => string.Equals(r.Field<string>("Status"),
                                                          "Active",
                                                          StringComparison.OrdinalIgnoreCase));

        // If no rows match, create an empty table with the same schema to avoid exceptions.
        DataTable filteredTable = activeRows.Any()
            ? activeRows.CopyToDataTable()
            : csvTable.Clone(); // empty table with same columns

        // --------------------------------------------------------------------
        // Build the report using the filtered data.
        // --------------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // The data source name ("persons") must match the name used in the template tags.
        engine.BuildReport(doc, filteredTable, "persons");

        // --------------------------------------------------------------------
        // Save the rendered document.
        // --------------------------------------------------------------------
        doc.Save("FilteredReport.docx");
    }
}
