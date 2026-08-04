using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Preserve the original thread culture.
        CultureInfo originalCulture = System.Threading.Thread.CurrentThread.CurrentCulture;

        // Set a custom culture (French - France) that uses a comma as the decimal separator.
        System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("fr-FR");

        // -------------------------------------------------
        // Create a template document with a LINQ Reporting tag.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // The tag references the "Total" property of the root object named "model".
        builder.Writeln("Total amount: <<[model.Total]>>");

        // -------------------------------------------------
        // Prepare the data model.
        // -------------------------------------------------
        ReportModel model = new ReportModel { Total = 12345.67m };

        // -------------------------------------------------
        // Build the report using the ReportingEngine.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        engine.BuildReport(doc, model, "model");

        // -------------------------------------------------
        // Save the generated report.
        // -------------------------------------------------
        doc.Save("ReportOutput.docx");

        // Restore the original culture.
        System.Threading.Thread.CurrentThread.CurrentCulture = originalCulture;
    }

    // Simple data model with a numeric property.
    public class ReportModel
    {
        public decimal Total { get; set; } = 0;
    }
}
