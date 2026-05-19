using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LibraryA
{
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }
}

namespace LibraryB
{
    public class Address
    {
        public string City { get; set; } = "";
        public string Country { get; set; } = "";
    }
}

public class ReportModel
{
    public LibraryA.Person Person { get; set; } = new();
    public LibraryB.Address Address { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a template document with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Name: <<[model.Person.Name]>>");
        builder.Writeln("Age: <<[model.Person.Age]>>");
        builder.Writeln("City: <<[model.Address.City]>>");
        builder.Writeln("Country: <<[model.Address.Country]>>");

        // Save the template to disk.
        string templatePath = Path.Combine(outputDir, "template.docx");
        template.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Prepare the data model.
        ReportModel model = new ReportModel
        {
            Person = new LibraryA.Person { Name = "John Doe", Age = 30 },
            Address = new LibraryB.Address { City = "New York", Country = "USA" }
        };

        // Register external types from multiple assemblies.
        ReportingEngine engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(LibraryA.Person));
        engine.KnownTypes.Add(typeof(LibraryB.Address));

        // Build the report using the model and the root name "model".
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(outputDir, "report.docx");
        doc.Save(outputPath);
    }
}
