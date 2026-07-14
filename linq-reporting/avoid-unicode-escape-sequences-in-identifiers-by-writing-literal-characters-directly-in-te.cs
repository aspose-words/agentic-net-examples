using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for Aspose.Words).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data with a Unicode property name.
        var persons = new List<Person>
        {
            new Person { 名 = "山田太郎", Age = 30 },
            new Person { 名 = "李小龍", Age = 35 },
            new Person { 名 = "José", Age = 28 }
        };

        var model = new ReportModel { Persons = persons };

        // Create a template document programmatically.
        var templatePath = "Template.docx";
        var builder = new DocumentBuilder();
        builder.Writeln("<<foreach [person in Persons]>>");
        builder.Writeln("Name: <<[person.名]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");
        builder.Document.Save(templatePath);

        // Load the template and build the report.
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        var outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}

// Wrapper class that matches the root object name used in BuildReport.
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

// Data model with a Unicode identifier used directly in the template.
public class Person
{
    public string 名 { get; set; } = string.Empty;
    public int Age { get; set; }
}
