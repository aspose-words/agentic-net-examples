using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

#nullable enable

public class Program
{
    public static void Main()
    {
        // Create output folder.
        string folder = "Output";
        Directory.CreateDirectory(folder);

        // 1. Create the template document.
        string templatePath = Path.Combine(folder, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Persons Report:");
        // LINQ Reporting foreach tag.
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // 2. Load the template for reporting.
        Document doc = new Document(templatePath);

        // 3. Prepare the data model with an empty collection (treated as absent enumeration).
        ReportModel model = new()
        {
            Persons = new() // Empty list instead of null to avoid exceptions.
        };

        // 4. Create the reporting engine (no special options needed for this scenario).
        ReportingEngine engine = new();

        // 5. Build the report.
        engine.BuildReport(doc, model, "model");

        // 6. Save the generated report.
        string outputPath = Path.Combine(folder, "Report.docx");
        doc.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
    }
}

// Data model classes.
public class ReportModel
{
    // When empty, the foreach loop produces no output.
    public List<Person> Persons { get; set; } = new();
}

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}
