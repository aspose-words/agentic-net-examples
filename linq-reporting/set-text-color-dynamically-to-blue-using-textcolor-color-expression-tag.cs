using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare folders
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "template.docx");
        string outputPath = Path.Combine(workDir, "output.docx");

        // Create the template document with a textColor tag
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("<<textColor [model.ColorName]>>This text is blue<</textColor>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting
        Document doc = new Document(templatePath);

        // Prepare the data model
        ReportModel model = new ReportModel
        {
            ColorName = "Blue"
        };

        // Build the report
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated document
        doc.Save(outputPath);
    }
}

// Data model used by the LINQ Reporting engine
public class ReportModel
{
    // The color expression used by the <<textColor>> tag
    public string ColorName { get; set; } = "Blue";
}
