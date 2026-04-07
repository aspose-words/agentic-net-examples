using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create output folder
        string outputFolder = "Output";
        Directory.CreateDirectory(outputFolder);

        // Create a template document with a placeholder for the total amount
        string templatePath = Path.Combine(outputFolder, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("The total amount is: <<[model.Total]>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting
        Document doc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Prepare the data model
        ReportModel model = new ReportModel { Total = 12345 };

        // Build the report using the model as the root object named "model"
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        string resultPath = Path.Combine(outputFolder, "Report.docx");
        doc.Save(resultPath);
    }

    public class ReportModel
    {
        public int Total { get; set; }
    }
}
