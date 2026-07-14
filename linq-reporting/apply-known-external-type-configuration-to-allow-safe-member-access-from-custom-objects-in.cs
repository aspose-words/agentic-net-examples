using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Create a template document with LINQ Reporting tags
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Current date (raw): <<[Date]>>");
        builder.Writeln("Formatted date (custom helper): <<[MyHelper.FormatDate(Date)]>>");
        builder.Writeln("Pi value (Math): <<[Math.PI]>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting
        Document doc = new Document(templatePath);

        // Prepare the data model
        ReportModel model = new ReportModel
        {
            Date = DateTime.Now
        };

        // Configure the ReportingEngine and register known types
        ReportingEngine engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(MyHelper));
        engine.KnownTypes.Add(typeof(Math));

        // Build the report using the model as the root object named "model"
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}

// Public data model class
public class ReportModel
{
    public DateTime Date { get; set; } = DateTime.Now;
}

// Static helper class with a method accessible from the template
public static class MyHelper
{
    public static string FormatDate(DateTime dt)
    {
        return dt.ToString("yyyy-MM-dd");
    }
}
