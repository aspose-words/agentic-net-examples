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
        // Ensure CSV file exists with sample data.
        const string csvPath = "People.csv";
        if (!File.Exists(csvPath))
        {
            File.WriteAllLines(csvPath, new[]
            {
                "Name,Age,Status",
                "John,30,Active",
                "Jane,25,Inactive",
                "Bob,40,Active"
            });
        }

        // Create a simple Word template in memory if the file does not exist.
        const string templatePath = "Template.docx";
        Document doc;
        if (File.Exists(templatePath))
        {
            doc = new Document(templatePath);
        }
        else
        {
            doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("Report of active persons:");
            builder.Writeln("<<foreach [persons]>>");
            builder.Writeln("Name: <<[Name]>>, Age: <<[Age]>>");
            builder.Writeln("<</foreach>>");
        }

        // Load CSV data into a DataTable.
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
                    foreach (var header in fields)
                        csvTable.Columns.Add(header.Trim());
                    isFirstLine = false;
                }
                else
                {
                    var row = csvTable.NewRow();
                    for (int i = 0; i < fields.Length; i++)
                        row[i] = fields[i].Trim();
                    csvTable.Rows.Add(row);
                }
            }
        }

        // Filter rows where Status == "Active".
        var activeRows = csvTable.AsEnumerable()
                                 .Where(r => string.Equals(r.Field<string>("Status"),
                                                          "Active",
                                                          StringComparison.OrdinalIgnoreCase));

        DataTable filteredTable = activeRows.Any()
            ? activeRows.CopyToDataTable()
            : csvTable.Clone();

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, filteredTable, "persons");

        // Save the rendered document.
        doc.Save("FilteredReport.docx");
    }
}
