using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

public partial class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report
        string templatePath = "Template.docx";
        string reportPath = "Report.docx";

        // -------------------------------------------------
        // Create the LINQ Reporting template programmatically
        // -------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Persons Report");
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Load the template for report generation
        // -------------------------------------------------
        var doc = new Document(templatePath);

        // -------------------------------------------------
        // Prepare sample data
        // -------------------------------------------------
        var model = new ReportModel();
        model.Persons.Add(new Person { Name = "Alice", Age = 30 });
        model.Persons.Add(new Person { Name = "Bob", Age = 25 });
        model.Persons.Add(new Person { Name = "Charlie", Age = 28 });

        // Demonstrate a foreach loop without explicit type (using var)
        foreach (var person in model.Persons)
        {
            Console.WriteLine($"{person.Name} is {person.Age} years old.");
        }

        // -------------------------------------------------
        // Build the report using Aspose.Words LINQ Reporting Engine
        // -------------------------------------------------
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        doc.Save(reportPath);
    }
}
