using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV handling.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data.
        string csvPath = "data.csv";
        File.WriteAllLines(csvPath, new[]
        {
            "Id,Name,Status",
            "1,Alpha,Active",
            "2,Beta,Inactive",
            "3,Gamma,Active",
            "4,Delta,Inactive"
        });

        // Load CSV into a DataTable.
        DataTable table = new DataTable();
        using (var reader = new StreamReader(csvPath))
        {
            bool isFirstLine = true;
            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine();
                if (string.IsNullOrWhiteSpace(line)) continue;
                var parts = line.Split(',');

                if (isFirstLine)
                {
                    foreach (var col in parts)
                        table.Columns.Add(col.Trim());
                    isFirstLine = false;
                }
                else
                {
                    table.Rows.Add(parts);
                }
            }
        }

        // Filter rows where Status == "Active".
        var activeItems = table.AsEnumerable()
            .Where(r => r.Field<string>("Status") == "Active")
            .Select(r => new Person
            {
                Id = r.Field<string>("Id"),
                Name = r.Field<string>("Name"),
                Status = r.Field<string>("Status")
            })
            .ToList();

        // Wrap filtered data in a model class.
        var model = new ReportModel { Items = activeItems };

        // Create a template document with LINQ Reporting tags.
        string templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Report of Active Items:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Id: <<[item.Id]>>, Name: <<[item.Name]>>, Status: <<[item.Status]>>");
        builder.Writeln("<</foreach>>");
        doc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Build the report using the filtered model.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        string outputPath = "report.docx";
        reportDoc.Save(outputPath);
    }
}

// Model representing a single CSV row.
public class Person
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public string Status { get; set; } = "";
}

// Wrapper class used as the root data source for the report.
public class ReportModel
{
    public List<Person> Items { get; set; } = new();
}
