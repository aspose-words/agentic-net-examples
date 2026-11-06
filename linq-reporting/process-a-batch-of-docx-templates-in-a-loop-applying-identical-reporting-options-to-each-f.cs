using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

public partial class Program
{
    // Sample data model
    public class ReportModel
    {
        public string Title { get; set; } = "Sample Report";
        public List<Person> Persons { get; set; } = new();
    }

    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    public static void Main()
    {
        // Prepare sample data
        var model = new ReportModel
        {
            Persons = new()
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 25 },
                new Person { Name = "Charlie", Age = 35 }
            }
        };

        // Create folders for templates and generated reports
        string templatesFolder = Path.Combine(Directory.GetCurrentDirectory(), "Templates");
        string reportsFolder = Path.Combine(Directory.GetCurrentDirectory(), "Reports");
        Directory.CreateDirectory(templatesFolder);
        Directory.CreateDirectory(reportsFolder);

        // Create a few identical template files
        for (int i = 1; i <= 3; i++)
        {
            string templatePath = Path.Combine(templatesFolder, $"Template{i}.docx");
            CreateTemplate(templatePath);
        }

        // Process each template in the folder
        foreach (string templateFile in Directory.GetFiles(templatesFolder, "*.docx"))
        {
            // Load the template document
            var doc = new Document(templateFile);

            // Configure the reporting engine
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // identical options for all files

            // Build the report using the same model and root name "model"
            engine.BuildReport(doc, model, "model");

            // Save the generated report
            string outputFileName = Path.GetFileNameWithoutExtension(templateFile) + "_Report.docx";
            string outputPath = Path.Combine(reportsFolder, outputFileName);
            doc.Save(outputPath);
        }
    }

    // Creates a simple LINQ Reporting template and saves it to the specified path
    private static void CreateTemplate(string path)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Title paragraph
        builder.Writeln("<<[Title]>>");
        builder.Writeln();

        // Foreach loop over Persons collection
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("- <<[p.Name]>> is <<[p.Age]>> years old.");
        builder.Writeln("<</foreach>>");

        // Save the template
        doc.Save(path);
    }
}
