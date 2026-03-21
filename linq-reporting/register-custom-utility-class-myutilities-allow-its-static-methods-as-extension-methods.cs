using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class MyUtilities
{
    public static string ToUpper(string input) => input?.ToUpper();
}

public class DataSource
{
    public string Name { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Create an empty document (or load a template if you have one)
        Document doc = new Document();

        // Create a ReportingEngine instance
        ReportingEngine engine = new ReportingEngine();

        // Register the custom utility class so its static members are available in templates
        engine.KnownTypes.Add(typeof(MyUtilities));

        // Example data source for the template
        var data = new DataSource { Name = "Aspose" };

        // Populate the template with data
        engine.BuildReport(doc, data);

        // Save the generated report
        doc.Save("Report.docx");
    }
}
