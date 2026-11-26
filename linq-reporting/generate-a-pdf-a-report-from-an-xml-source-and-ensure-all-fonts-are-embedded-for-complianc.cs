using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}

public class PeopleModel
{
    public List<Person> Person { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Working directory
        string workDir = Directory.GetCurrentDirectory();

        // 1. Prepare sample data
        var model = new PeopleModel
        {
            Person = new List<Person>
            {
                new Person { Name = "John Doe", Age = 30 },
                new Person { Name = "Jane Smith", Age = 25 }
            }
        };

        // 2. Create a Word template with LINQ Reporting tags
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("People Report");
        builder.Writeln("==============");
        builder.Writeln();
        builder.Writeln("<<foreach [p in Person]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Age: <<[p.Age]>>");
        builder.Writeln();
        builder.Writeln("<</foreach>>");

        string templatePath = Path.Combine(workDir, "Template.docx");
        template.Save(templatePath);

        // 3. Load the template for reporting
        Document reportDoc = new Document(templatePath);

        // 4. Build the report using the LINQ Reporting engine
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // 5. Save the report as PDF/A with all fonts embedded
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b,
            EmbedFullFonts = true
        };

        string pdfPath = Path.Combine(workDir, "PeopleReport.pdf");
        reportDoc.Save(pdfPath, pdfOptions);

        // Indicate completion
        Console.WriteLine($"Report generated: {pdfPath}");
    }
}
