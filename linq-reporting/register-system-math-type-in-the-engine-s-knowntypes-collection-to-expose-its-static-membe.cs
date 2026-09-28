using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    // No properties needed for this example.
}

public class Program
{
    public static void Main()
    {
        // Define file paths.
        string templatePath = "template.docx";
        string outputPath = "output.docx";

        // Create a new document and add a LINQ Reporting tag that uses System.Math.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Rounded value of Pi (2 decimals): <<[Math.Round(3.14159, 2)]>>");
        templateDoc.Save(templatePath);

        // Load the template document.
        Document doc = new Document(templatePath);

        // Create the reporting engine and register System.Math as a known type.
        ReportingEngine engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(System.Math));

        // Build the report using an empty model as the root object.
        Model model = new Model();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(outputPath);
    }
}
