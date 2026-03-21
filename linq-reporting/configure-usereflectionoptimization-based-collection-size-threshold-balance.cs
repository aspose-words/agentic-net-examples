using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Sample data collection whose size will determine the optimization setting.
        List<string> data = GetSampleData();

        // Threshold for switching the reflection optimization on/off.
        // Large collections benefit from the dynamic class generation (true),
        // while small collections avoid its overhead (false).
        int threshold = 1000;
        ReportingEngine.UseReflectionOptimization = data.Count > threshold;

        // Create a new document (could also be loaded from a template file).
        Document doc = new Document();

        // Build a report using the ReportingEngine.
        // Pass the collection directly with a name that the template would reference.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, data, "Items");

        // Save the generated document.
        doc.Save("Report.docx");
    }

    // Generates a sample list of strings. Adjust the loop count to test different thresholds.
    private static List<string> GetSampleData()
    {
        var list = new List<string>();
        for (int i = 0; i < 500; i++) // Example size; change to >1000 to enable optimization.
        {
            list.Add($"Item {i}");
        }
        return list;
    }
}
