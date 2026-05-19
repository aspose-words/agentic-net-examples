using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for legacy encodings (required by Aspose.Words on .NET Core).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data files.
        string dataFolder = Path.Combine(Environment.CurrentDirectory, "Data");
        Directory.CreateDirectory(dataFolder);

        string xmlPath = Path.Combine(dataFolder, "departments.xml");
        File.WriteAllText(xmlPath,
@"<Departments>
    <Department>
        <Name>Human Resources</Name>
    </Department>
    <Department>
        <Name>Research &amp; Development</Name>
    </Department>
</Departments>");

        string jsonPath = Path.Combine(dataFolder, "employees.json");
        File.WriteAllText(jsonPath,
@"[
    { ""Name"": ""Alice"", ""Age"": 30 },
    { ""Name"": ""Bob"", ""Age"": 45 },
    { ""Name"": ""Charlie"", ""Age"": 28 }
]");

        // Create the LINQ Reporting template programmatically.
        string templatePath = Path.Combine(dataFolder, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Outer loop over XML departments.
        builder.Writeln("<<foreach [dept in xml]>>");
        builder.Writeln("Department: <<[dept.Name]>>");
        builder.Writeln("Employees:");

        // Inner loop over JSON employees.
        builder.Writeln("<<foreach [emp in json]>>");
        builder.Writeln("- <<[emp.Name]>> (Age: <<[emp.Age]>>)");
        builder.Writeln("<</foreach>>"); // End inner foreach

        builder.Writeln("<</foreach>>"); // End outer foreach

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Create data source objects.
        var xmlDataSource = new XmlDataSource(xmlPath);
        var jsonDataSource = new JsonDataSource(jsonPath);

        // Build the report using both data sources.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;

        bool success = engine.BuildReport(
            reportDoc,
            new object[] { xmlDataSource, jsonDataSource },
            new string[] { "xml", "json" });

        // Save the generated report.
        string outputPath = Path.Combine(dataFolder, "report.docx");
        reportDoc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to: {outputPath}");
    }
}
