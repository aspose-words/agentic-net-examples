using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary template and the final report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // ---------- Create the LINQ Reporting template ----------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Add a heading.
        builder.Writeln("Persons Report");
        builder.Writeln();

        // LINQ Reporting foreach block.
        // The engine will repeat the inner paragraph for each item in the collection.
        builder.Writeln("<<foreach [p in Persons]>>");
        // This line may become empty if Name is empty, demonstrating paragraph removal.
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // ---------- Load the template ----------
        var doc = new Document(templatePath);

        // ---------- Prepare sample data ----------
        var model = new ReportModel
        {
            Persons = new()
            {
                new Person { Name = "John Doe", Age = 30 },
                new Person { Name = "", Age = 0 },          // This will produce empty paragraphs.
                new Person { Name = "Alice Smith", Age = 25 }
            }
        };

        // ---------- Build the report ----------
        var engine = new ReportingEngine
        {
            // Remove paragraphs that become empty after processing the tags.
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };
        engine.BuildReport(doc, model, "model");

        // ---------- Save the final document ----------
        doc.Save(reportPath);
    }
}

// Root data model referenced by the template as <<[model.Persons]>>.
public class ReportModel
{
    // Initialize to avoid nullable warnings.
    public List<Person> Persons { get; set; } = new();
}

// Simple item class used inside the foreach loop.
public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}
