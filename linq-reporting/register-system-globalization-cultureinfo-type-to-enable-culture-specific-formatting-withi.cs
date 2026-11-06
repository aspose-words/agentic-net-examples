using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Set French culture for formatting
        CultureInfo.CurrentCulture = new CultureInfo("fr-FR");
        CultureInfo.CurrentUICulture = new CultureInfo("fr-FR");

        // Ensure output directory exists
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a template document with LINQ Reporting tags
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Customer: <<[model.CustomerName]>>");
        builder.Writeln("Date (French format): <<[model.OrderDate.ToString(\"d\")]>>");
        builder.Writeln("Amount (French format): <<[model.Amount.ToString(\"C\")]>>");

        templateDoc.Save(templatePath);

        // Load the template
        Document doc = new Document(templatePath);

        // Prepare the data model
        OrderReport model = new OrderReport
        {
            CustomerName = "Jean Dupont",
            OrderDate = new DateTime(2023, 12, 15),
            Amount = 1234.56m
        };

        // Configure the reporting engine
        ReportingEngine engine = new ReportingEngine();

        // Build the report
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        string reportPath = Path.Combine(outputDir, "report.docx");
        doc.Save(reportPath);

        Console.WriteLine($"Report generated at: {reportPath}");
    }
}

// Data model class
public class OrderReport
{
    public string CustomerName { get; set; } = "";
    public DateTime OrderDate { get; set; }
    public decimal Amount { get; set; }
}
