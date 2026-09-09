using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Sample data property.
    public string Name { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder to insert the LINQ Reporting template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Template:
        // - The first character of the Name property is wrapped in a textColor tag (red).
        // - The remaining characters are output without additional formatting.
        builder.Writeln(
            "<<textColor [\"Red\"]>><<[model.Name.Substring(0,1)]>><</textColor>><<[model.Name.Substring(1)]>>");

        // Prepare the data source.
        ReportModel model = new ReportModel { Name = "Aspose.Words" };

        // Build the report using the model as the root object named "model".
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the resulting document.
        doc.Save("CustomFontStyling.docx");
    }
}
