using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Enable code pages for any required encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // File paths.
        string xmlPath = "departments.xml";
        string jsonPath = "employees.json";
        string templatePath = "template.docx";
        string outputPath = "report.docx";

        // Create sample XML data source.
        File.WriteAllText(xmlPath,
@"<Departments>
    <Department>
        <Id>1</Id>
        <Name>Human Resources</Name>
    </Department>
    <Department>
        <Id>2</Id>
        <Name>Information Technology</Name>
    </Department>
    <Department>
        <Id>3</Id>
        <Name>Finance</Name>
    </Department>
</Departments>");

        // Create sample JSON data source.
        File.WriteAllText(jsonPath,
@"[
    { ""Id"": 1, ""Name"": ""Alice"", ""Title"": ""HR Manager"", ""DepartmentId"": 1 },
    { ""Id"": 2, ""Name"": ""Bob"", ""Title"": ""Recruiter"", ""DepartmentId"": 1 },
    { ""Id"": 3, ""Name"": ""Charlie"", ""Title"": ""Developer"", ""DepartmentId"": 2 },
    { ""Id"": 4, ""Name"": ""Diana"", ""Title"": ""System Analyst"", ""DepartmentId"": 2 },
    { ""Id"": 5, ""Name"": ""Eve"", ""Title"": ""Accountant"", ""DepartmentId"": 3 }
]");

        // Build the template document with nested foreach tags.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Company Report");
        builder.Writeln("----------------");
        builder.Writeln();

        // Outer loop over XML departments (named "xml").
        builder.Writeln("<<foreach [dept in xml]>>");
        builder.Writeln("Department: <<[dept.Name]>>");
        builder.Writeln();

        // Inner loop over JSON employees (named "json").
        builder.Writeln("<<foreach [emp in json]>>");
        builder.Writeln(" - <<[emp.Name]>> (<<[emp.Title]>>), DeptId: <<[emp.DepartmentId]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Create data source objects.
        var xmlData = new XmlDataSource(xmlPath);
        var jsonData = new JsonDataSource(jsonPath);

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Build the report using both data sources.
        engine.BuildReport(reportDoc,
            new object[] { xmlData, jsonData },
            new string[] { "xml", "json" });

        // Save the final report.
        reportDoc.Save(outputPath);
    }
}
