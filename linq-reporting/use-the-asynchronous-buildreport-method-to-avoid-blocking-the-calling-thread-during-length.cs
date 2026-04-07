using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Reporting;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}

public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

public class ReportGenerator
{
    // Entry point – async to allow awaiting the report generation.
    public static async Task Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 45 },
                new Person { Name = "Charlie", Age = 28 }
            }
        };

        // Create a template document programmatically.
        const string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Build the report asynchronously.
        await BuildReportAsync(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Report generated and saved to '{Path.GetFullPath(outputPath)}'.");
    }

    // Creates a simple Word template containing LINQ Reporting tags.
    private static void CreateTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Add a title.
        builder.Writeln("Person Report");
        builder.Writeln();

        // LINQ Reporting foreach loop over the collection.
        builder.Writeln("<<foreach [person in Persons]>>");
        builder.Writeln("- <<[person.Name]>> (Age: <<[person.Age]>>)");
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(filePath);
    }

    // Asynchronously builds the report using ReportingEngine.
    private static Task BuildReportAsync(Document doc, object dataSource, string dataSourceName)
    {
        return Task.Run(() =>
        {
            var engine = new ReportingEngine();
            // No special options required for this simple example.
            engine.BuildReport(doc, dataSource, dataSourceName);
        });
    }
}
