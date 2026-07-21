using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for legacy encodings (required by Aspose.Words in some environments)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data
        var person = new Person
        {
            Name = "Alice Johnson"
        };

        // Create a template document programmatically
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Write a line that formats the first character of the name in red,
        // then writes the rest of the name without special formatting.
        builder.Writeln("Name: <<textColor [\"Red\"]>><<[person.Name.Substring(0,1)]>><</textColor>><<[person.Name.Substring(1)]>>");

        doc.Save(templatePath);

        // Load the template for reporting
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the person object as the root data source named "person"
        engine.BuildReport(reportDoc, person, "person");

        // Save the generated report
        var outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

// Simple data model used by the LINQ Reporting engine
public class Person
{
    public string Name { get; set; } = string.Empty;
}
