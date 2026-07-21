using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Record
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
    public string City { get; set; } = "";
}

public class ReportModel
{
    public List<Record> FilteredRecords { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create sample JSON data.
        string jsonPath = "data.json";
        var sampleData = new[]
        {
            new Record { Name = "John Doe", Age = 45, City = "New York" },
            new Record { Name = "Jane Smith", Age = 28, City = "Los Angeles" },
            new Record { Name = "Jack Johnson", Age = 38, City = "New York" },
            new Record { Name = "Emily Davis", Age = 50, City = "Chicago" },
            new Record { Name = "James Brown", Age = 33, City = "New York" }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

        // Load JSON records.
        var records = JsonConvert.DeserializeObject<List<Record>>(File.ReadAllText(jsonPath)) ?? new List<Record>();

        // Apply complex predicate using Where.
        var filtered = records
            .Where(r => r.Age > 30 && r.City == "New York" && r.Name.StartsWith("J"))
            .ToList();

        // Prepare the model for reporting.
        var model = new ReportModel { FilteredRecords = filtered };

        // Create a Word template programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Filtered Records Report");
        builder.Writeln("<<foreach [rec in FilteredRecords]>>");
        builder.Writeln("Name: <<[rec.Name]>>, Age: <<[rec.Age]>>, City: <<[rec.City]>>");
        builder.Writeln("<</foreach>>");

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
