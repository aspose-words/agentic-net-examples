using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace ModelsA
{
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }
}

namespace ModelsB
{
    public class Address
    {
        public string City { get; set; } = "";
        public string Country { get; set; } = "";
    }
}

public class ReportData
{
    public ModelsA.Person Person { get; set; } = new();
    public ModelsB.Address Address { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Create template document with LINQ Reporting tags.
        const string templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Customer Report");
        builder.Writeln("Name: <<[data.Person.Name]>>");
        builder.Writeln("Age: <<[data.Person.Age]>>");
        builder.Writeln("City: <<[data.Address.City]>>");
        builder.Writeln("Country: <<[data.Address.Country]>>");
        doc.Save(templatePath);

        // Load the template.
        var template = new Document(templatePath);

        // Prepare data.
        var data = new ReportData
        {
            Person = new ModelsA.Person { Name = "John Doe", Age = 30 },
            Address = new ModelsB.Address { City = "New York", Country = "USA" }
        };

        // Configure ReportingEngine.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;

        // Build the report.
        bool success = engine.BuildReport(template, data, "data");

        // Save the generated report.
        const string outputPath = "Report.docx";
        template.Save(outputPath);

        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to {Path.GetFullPath(outputPath)}");
    }
}
