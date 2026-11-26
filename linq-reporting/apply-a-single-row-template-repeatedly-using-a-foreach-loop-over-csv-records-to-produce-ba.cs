using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Record
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
    public string City { get; set; } = "";
}

public class ReportModel
{
    public List<Record> Records { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Create sample CSV data.
        string csvPath = "data.csv";
        File.WriteAllText(csvPath,
            "Name,Age,City\n" +
            "Alice,30,New York\n" +
            "Bob,25,Los Angeles\n" +
            "Charlie,35,Chicago");

        // Load CSV into the model.
        ReportModel model = new();
        string[] lines = File.ReadAllLines(csvPath);
        for (int i = 1; i < lines.Length; i++)
        {
            string[] parts = lines[i].Split(',');
            if (parts.Length >= 3)
            {
                model.Records.Add(new Record
                {
                    Name = parts[0],
                    Age = int.TryParse(parts[1], out int age) ? age : 0,
                    City = parts[2]
                });
            }
        }

        // Build the template document programmatically.
        string templatePath = "template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin the foreach loop that will repeat the whole table for each record.
        builder.Writeln("<<foreach [rec in Records]>>");

        // Title for each record (optional).
        builder.Writeln("Customer Report");

        // Table for a single record.
        builder.StartTable();

        // Header row.
        builder.InsertCell(); builder.Writeln("Name");
        builder.InsertCell(); builder.Writeln("Age");
        builder.InsertCell(); builder.Writeln("City");
        builder.EndRow();

        // Data row – placeholders will be replaced by the engine.
        builder.InsertCell(); builder.Writeln("<<[rec.Name]>>");
        builder.InsertCell(); builder.Writeln("<<[rec.Age]>>");
        builder.InsertCell(); builder.Writeln("<<[rec.City]>>");
        builder.EndRow();

        // End of the table for this record.
        builder.EndTable();

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the final document.
        string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
