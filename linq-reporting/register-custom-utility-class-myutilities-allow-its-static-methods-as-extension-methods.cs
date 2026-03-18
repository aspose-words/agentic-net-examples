using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class MyUtilities
{
    // Example static method that can be called from a reporting template
    public static string ToUpper(string input) => input?.ToUpper();
}

public class Program
{
    public static void Main()
    {
        // Load the template document (adjust the path as needed)
        Document doc = new Document("Template.docx");

        // Create a ReportingEngine instance
        ReportingEngine engine = new ReportingEngine();

        // Register the custom utility class so its static members are available in templates
        engine.KnownTypes.Add(typeof(MyUtilities));

        // Example data source for the template
        var data = new { Name = "Aspose" };

        // Populate the template with data
        engine.BuildReport(doc, data);

        // Save the generated report
        doc.Save("Report.docx");
    }
}
