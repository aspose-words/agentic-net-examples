using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = string.Empty;
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
        const string templatePath = "template.docx";
        const string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report of Persons:");
        builder.Writeln("<<foreach [p in Persons]>>");
        // This paragraph may become empty if p.Name is empty.
        builder.Writeln("<<[p.Name]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template for report generation.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Sample data: one empty name to produce an empty paragraph.
        var model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice" },
                new Person { Name = "" },      // Will generate an empty paragraph.
                new Person { Name = "Bob" }
            }
        };

        // -----------------------------------------------------------------
        // Configure the ReportingEngine to remove empty paragraphs.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Build the report.
        engine.BuildReport(doc, model, "model");

        // Save the final document.
        doc.Save(outputPath);
    }
}
