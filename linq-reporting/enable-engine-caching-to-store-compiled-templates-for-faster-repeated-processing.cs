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
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var model = new ReportModel
        {
            Persons = new()
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 25 },
                new Person { Name = "Charlie", Age = 35 }
            }
        };

        // Create a template document programmatically.
        const string templatePath = "template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        var templateDoc = new Document(templatePath);

        // Create a ReportingEngine. Caching of compiled templates is enabled by default.
        var engine = new ReportingEngine();

        // First report generation (template will be compiled and cached).
        const string outputPath1 = "output1.docx";
        engine.BuildReport(templateDoc, model, "model");
        templateDoc.Save(outputPath1);

        // Load the template again for a second run.
        var templateDoc2 = new Document(templatePath);

        // Second report generation (should use cached compiled template).
        const string outputPath2 = "output2.docx";
        engine.BuildReport(templateDoc2, model, "model");
        templateDoc2.Save(outputPath2);
    }

    private static void CreateTemplate(string path)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Person Report");
        builder.Writeln("==============");
        builder.Writeln();
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(path);
    }
}

public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
