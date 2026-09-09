using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class KnownTypesExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a LINQ Reporting tag that accesses a static member of DateTime.
        // The tag will be resolved only if DateTime is added to the engine's KnownTypes collection.
        builder.Writeln("Current date and time: <<[DateTime.Now]>>");

        // Prepare the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Register the external type (DateTime) so its static members can be used safely in the template.
        engine.KnownTypes.Add(typeof(DateTime));

        // No data source is required for this simple example, but we still need to call BuildReport.
        // Pass an empty object as the data source and an empty name because the template does not reference a root object.
        engine.BuildReport(doc, new object(), "");

        // Save the generated report.
        doc.Save("KnownTypesReport.docx");
    }
}
