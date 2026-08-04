using System;
using System.Collections.Generic;
using System.Diagnostics;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "John Doe", Age = 30, MiddleName = "A." },
                new Person { Name = "Jane Smith", Age = 25, MiddleName = "" },
                new Person { Name = "Bob Johnson", Age = 40, MiddleName = null! } // Will be treated as empty.
            }
        };

        // Create a template document with LINQ Reporting tags.
        var templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Measure memory before report generation.
        long memoryBefore = GC.GetTotalMemory(true);

        // Build the report with RemoveEmptyParagraphs option enabled.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };
        engine.BuildReport(doc, model, "model");

        // Measure memory after report generation.
        long memoryAfter = GC.GetTotalMemory(true);

        // Output memory consumption information.
        Console.WriteLine($"Memory before report: {memoryBefore:N0} bytes");
        Console.WriteLine($"Memory after report : {memoryAfter:N0} bytes");
        Console.WriteLine($"Memory increase      : {memoryAfter - memoryBefore:N0} bytes");

        // Save the generated report.
        doc.Save("Report.docx");
    }

    private static void CreateTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Header paragraph.
        builder.Writeln("People Report");
        builder.Writeln();

        // Begin foreach loop over Persons collection.
        builder.Writeln("<<foreach [p in Persons]>>");

        // Output name.
        builder.Writeln("Name: <<[p.Name]>>");

        // Conditional output of middle name; paragraph may become empty.
        builder.Writeln("<<if [p.MiddleName != \"\"]>>Middle: <<[p.MiddleName]>> <</if>>");

        // Output age.
        builder.Writeln("Age: <<[p.Age]>>");

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(filePath);
    }
}

// Wrapper class for the data source.
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

// Data model class.
public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
    public string MiddleName { get; set; } = "";
}
