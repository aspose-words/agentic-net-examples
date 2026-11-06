using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class ReportModel
{
    public List<string> FormattedPersons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample data.
        List<Person> persons = new()
        {
            new Person { Name = "Alice", Age = 30 },
            new Person { Name = "Bob", Age = 25 },
            new Person { Name = "Charlie", Age = 35 }
        };

        // Transform each person into a formatted string using a custom delegate.
        ReportModel model = new()
        {
            FormattedPersons = persons
                .Select(p => $"{p.Name} ({p.Age} years old)")
                .ToList()
        };

        // Create the LINQ Reporting template.
        string templatePath = "Template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        // Write heading.
        builder.Writeln("Person List:");
        // Begin foreach over formatted strings.
        builder.Writeln("<<foreach [p in FormattedPersons]>>");
        // Output each formatted string.
        builder.Writeln("<<[p]>>");
        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new(templatePath);

        // Build the report.
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the final report.
        string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
