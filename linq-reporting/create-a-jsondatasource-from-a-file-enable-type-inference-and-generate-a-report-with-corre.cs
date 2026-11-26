using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required by Aspose.Words for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data model
        var customer = new Customer
        {
            Name = "John Doe",
            Age = 30,
            IsMember = true,
            Balance = 1234.56m
        };

        // Serialize model to JSON and write to a file
        string jsonPath = "data.json";
        string jsonContent = JsonConvert.SerializeObject(customer, Formatting.Indented);
        File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

        // Create a Word template programmatically
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Customer Report");
        builder.Writeln("Name: <<[root.Name]>>");
        builder.Writeln("Age: <<[root.Age]>>");
        builder.Writeln("Member: <<[root.IsMember]>>");
        builder.Writeln("Balance: $<<[root.Balance]>>");

        string templatePath = "template.docx";
        templateDoc.Save(templatePath);

        // Load the template
        var doc = new Document(templatePath);

        // Configure JSON data source (type inference is enabled by default in recent versions)
        var jsonDataSource = new JsonDataSource(jsonPath);

        // Build the report
        var engine = new ReportingEngine();
        engine.BuildReport(doc, jsonDataSource, "root");

        // Save the generated report
        string outputPath = "report.docx";
        doc.Save(outputPath);
    }
}

// Data model used only for JSON generation; the report uses JsonDataSource directly.
public class Customer
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
    public bool IsMember { get; set; }
    public decimal Balance { get; set; }
}
