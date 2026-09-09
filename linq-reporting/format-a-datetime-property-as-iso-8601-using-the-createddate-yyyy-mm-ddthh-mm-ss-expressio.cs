using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Sample DateTime property; initialized to current time.
    public DateTime CreatedDate { get; set; } = DateTime.Now;

    // Returns the date formatted as ISO 8601.
    public string CreatedDateIso => CreatedDate.ToString("yyyy-MM-ddTHH:mm:ss");
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required on .NET Core).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a LINQ Reporting tag that uses the pre‑formatted property.
        builder.Writeln("Report generated at: <<[model.CreatedDateIso]>>");

        // Prepare the data source.
        ReportModel model = new ReportModel();

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the resulting document.
        doc.Save("ReportOutput.docx");
    }
}
