using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Quantity is a double to demonstrate the need for an explicit cast to int.
    public double Quantity { get; set; } = 0;

    // This method expects an integer parameter.
    public string GetMessage(int qty)
    {
        return $"The quantity (cast to int) is {qty}.";
    }
}

public class Program
{
    public static void Main()
    {
        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a LINQ Reporting tag that calls GetMessage with an explicit cast to int.
        builder.Writeln("<<[model.GetMessage((int)model.Quantity)]>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template back for reporting.
        Document doc = new Document(templatePath);

        // Prepare the data model.
        ReportModel model = new ReportModel
        {
            Quantity = 7.9 // Example value that will be cast to int (7).
        };

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string reportPath = "Report.docx";
        doc.Save(reportPath);
    }
}
