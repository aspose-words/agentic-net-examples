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
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for the template and the final report
        string templatePath = "Template.docx";
        string outputPath = "Report.docx";

        // -------------------------------------------------
        // Create the LINQ Reporting template programmatically
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title
        builder.Writeln("People Report");
        builder.Writeln();

        // Foreach block that will be populated from the data source
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // A paragraph that will become empty if the collection is empty
        builder.Writeln();
        builder.Writeln("End of Report");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Load the template and build the report
        // -------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Prepare data: first a populated list, then an empty list to demonstrate removal
        ReportModel model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 45 }
            }
        };

        // Configure the ReportingEngine to remove empty paragraphs after processing
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };

        // Build the report using the model; the root name in the template is "model"
        engine.BuildReport(reportDoc, model, "model");

        // Save the final document
        reportDoc.Save(outputPath);
    }
}

// -----------------------------------------------------------------
// Data model classes (public with initialized properties to avoid warnings)
// -----------------------------------------------------------------
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
