using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class Model
{
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Create a simple template with LINQ Reporting tags.
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Persons Report");
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("- <<[p.Name]>> (Age: <<[p.Age]>>)");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Restrict prohibited .NET types before building the report.
        ReportingEngine.SetRestrictedTypes(new Type[]
        {
            typeof(System.IO.FileInfo),
            typeof(System.Diagnostics.Process)
        });

        // Prepare sample data.
        Model model = new Model
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 25 }
            }
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
