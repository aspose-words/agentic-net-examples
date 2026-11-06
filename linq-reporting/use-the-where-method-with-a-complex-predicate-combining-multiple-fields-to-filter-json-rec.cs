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
    public bool IsActive { get; set; }
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

        // Prepare sample JSON data.
        string jsonPath = "data.json";
        var sampleData = new List<Record>
        {
            new Record { Name = "Alice", Age = 28, City = "London", IsActive = true },
            new Record { Name = "Bob", Age = 35, City = "London", IsActive = false },
            new Record { Name = "Charlie", Age = 42, City = "London", IsActive = true },
            new Record { Name = "Diana", Age = 31, City = "Paris", IsActive = true },
            new Record { Name = "Ethan", Age = 45, City = "London", IsActive = true }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

        // Load JSON data.
        var allRecords = JsonConvert.DeserializeObject<List<Record>>(File.ReadAllText(jsonPath)) ?? new List<Record>();

        // Apply complex predicate: Age > 30, City == "London", and IsActive == true.
        var filtered = allRecords
            .Where(r => r.Age > 30 && r.City == "London" && r.IsActive)
            .ToList();

        // Prepare the model for the report.
        var model = new ReportModel
        {
            FilteredRecords = filtered
        };

        // Create the template document programmatically.
        var templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Filtered Records (Age > 30, City = London, Active):");
        builder.Writeln("<<foreach [rec in FilteredRecords]>>");
        builder.Writeln("Name: <<[rec.Name]>>, Age: <<[rec.Age]>>, City: <<[rec.City]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template and build the report.
        var templateDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(templateDoc, model, "model");

        // Save the generated report.
        var outputPath = "report.docx";
        templateDoc.Save(outputPath);
    }
}
