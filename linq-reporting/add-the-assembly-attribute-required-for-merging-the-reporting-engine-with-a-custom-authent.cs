using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

[assembly: Aspose.Words.Reporting.CustomAuthenticationModule(typeof(MyAuthModule))]

namespace Aspose.Words.Reporting
{
    [AttributeUsage(AttributeTargets.Assembly)]
    public sealed class CustomAuthenticationModuleAttribute : Attribute
    {
        public Type ModuleType { get; }

        public CustomAuthenticationModuleAttribute(Type moduleType) => ModuleType = moduleType;
    }
}

public class MyAuthModule
{
    // Placeholder for custom authentication logic.
    // The real implementation would depend on the Aspose.Words.Reporting API.
}

public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}

public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var persons = new List<Person>
        {
            new() { Name = "Alice", Age = 30 },
            new() { Name = "Bob", Age = 25 },
            new() { Name = "Charlie", Age = 35 }
        };

        // Create a template document programmatically.
        var templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Persons Report");
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template and build the report.
        var template = new Document(templatePath);
        var engine = new ReportingEngine();

        var model = new ReportModel { Persons = persons };
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        var outputPath = "report.docx";
        template.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
