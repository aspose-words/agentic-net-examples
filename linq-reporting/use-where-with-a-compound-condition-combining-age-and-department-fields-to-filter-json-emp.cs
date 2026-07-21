using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider required by Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // ---------- Create sample JSON data ----------
        const string jsonFile = "employees.json";
        string json = @"[
  { ""Name"": ""Alice"",   ""Age"": 28, ""Department"": ""HR"" },
  { ""Name"": ""Bob"",     ""Age"": 35, ""Department"": ""HR"" },
  { ""Name"": ""Charlie"", ""Age"": 40, ""Department"": ""IT"" },
  { ""Name"": ""Diana"",   ""Age"": 32, ""Department"": ""HR"" }
]";
        File.WriteAllText(jsonFile, json);

        // ---------- Create the LINQ Reporting template ----------
        const string templateFile = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Filtered Employees (Age > 30 && Department == \"HR\"):");
        // Use Where with a compound condition inside the foreach tag.
        builder.Writeln("<<foreach [e in employees.Where(e => e.Age > 30 && e.Department == \"HR\")]>>");
        builder.Writeln("- <<[e.Name]>> (Age: <<[e.Age]>>, Dept: <<[e.Department]>>)");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templateFile);

        // ---------- Load the template ----------
        Document doc = new Document(templateFile);

        // ---------- Create JSON data source ----------
        JsonDataSource dataSource = new JsonDataSource(jsonFile);

        // ---------- Build the report ----------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "employees");

        // ---------- Save the generated report ----------
        const string outputFile = "Report.docx";
        doc.Save(outputFile);
    }
}
