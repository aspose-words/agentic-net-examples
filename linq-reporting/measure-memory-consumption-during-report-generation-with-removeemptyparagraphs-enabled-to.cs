using System;
using System.Collections.Generic;
using System.Diagnostics;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;
using System.Text;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
    public string? Optional { get; set; }
}

public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Paths for template and output documents.
        const string templatePath = "ReportTemplate.docx";
        const string outputPath = "ReportOutput.docx";

        // Create the template document with LINQ Reporting tags.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report Header");
        builder.Writeln("<<foreach [person in Persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        // This field may be empty; the paragraph should be removed when RemoveEmptyParagraphs is enabled.
        builder.Writeln("Optional: <<[person.Optional]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("Report Footer");
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document doc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30, Optional = "Info" },
                new Person { Name = "Bob", Age = 25, Optional = null },
                new Person { Name = "Charlie", Age = 40, Optional = "" }
            }
        };

        // Measure memory before building the report.
        long memoryBefore = GC.GetTotalMemory(true);

        // Configure the reporting engine with RemoveEmptyParagraphs option.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Build the report.
        engine.BuildReport(doc, model, "model");

        // Measure memory after building the report.
        long memoryAfter = GC.GetTotalMemory(true);

        // Save the generated report.
        doc.Save(outputPath);

        // Output memory consumption details.
        Console.WriteLine($"Memory before report generation: {memoryBefore} bytes");
        Console.WriteLine($"Memory after report generation:  {memoryAfter} bytes");
        Console.WriteLine($"Memory used by report generation: {memoryAfter - memoryBefore} bytes");
    }
}
