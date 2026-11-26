using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class ReportModel
{
    // All persons loaded from JSON.
    public List<Person> Persons { get; set; } = new();

    // Value to filter by (case‑insensitive).
    public string FilterName { get; set; } = "alice";

    // Persons whose Name matches FilterName ignoring case.
    public IEnumerable<Person> FilteredPersons => Persons
        .Where(p => p.Name.Equals(FilterName, StringComparison.OrdinalIgnoreCase));
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required by Aspose.Words for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // 1. Create sample JSON data file.
        string jsonPath = "people.json";
        var sampleData = new[]
        {
            new Person { Name = "Alice", Age = 30 },
            new Person { Name = "Bob", Age = 25 },
            new Person { Name = "alice", Age = 22 },
            new Person { Name = "Charlie", Age = 35 }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

        // 2. Load JSON into the model.
        var model = new ReportModel();
        string jsonContent = File.ReadAllText(jsonPath);
        model.Persons = JsonConvert.DeserializeObject<List<Person>>(jsonContent) ?? new List<Person>();

        // 3. Create the LINQ Reporting template programmatically.
        string templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Filtered Persons (Name = \"alice\", case‑insensitive):");
        builder.Writeln("<<foreach [p in FilteredPersons]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // 4. Load the template for reporting.
        var templateDoc = new Document(templatePath);

        // 5. Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(templateDoc, model, "model");

        // 6. Save the generated report.
        string outputPath = "report.docx";
        templateDoc.Save(outputPath);
    }
}
