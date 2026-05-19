using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some data sources)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Enable reflection optimization globally
        ReportingEngine.UseReflectionOptimization = true;

        // -----------------------------------------------------------------
        // 1. Create a simple LINQ Reporting template programmatically
        // -----------------------------------------------------------------
        string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare sample JSON data files
        // -----------------------------------------------------------------
        string largeJsonPath = "LargeData.json";
        string smallJsonPath = "SmallData.json";
        CreateLargeJson(largeJsonPath);
        CreateSmallJson(smallJsonPath);

        // -----------------------------------------------------------------
        // 3. Build report for large JSON array (optimization stays enabled)
        // -----------------------------------------------------------------
        Document largeReport = new Document(templatePath);
        JsonDataSource largeDataSource = new JsonDataSource(largeJsonPath);
        ReportingEngine largeEngine = new ReportingEngine();
        largeEngine.BuildReport(largeReport, largeDataSource, "persons");
        largeReport.Save("ReportLarge.docx");

        // -----------------------------------------------------------------
        // 4. Build report for small JSON array (disable optimization for this case)
        // -----------------------------------------------------------------
        ReportingEngine.UseReflectionOptimization = false; // selective disable
        Document smallReport = new Document(templatePath);
        JsonDataSource smallDataSource = new JsonDataSource(smallJsonPath);
        ReportingEngine smallEngine = new ReportingEngine();
        smallEngine.BuildReport(smallReport, smallDataSource, "persons");
        smallReport.Save("ReportSmall.docx");
    }

    // Creates a Word document containing a foreach loop over "persons"
    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // LINQ Reporting tags
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("<<[p.Name]>> - <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(filePath);
    }

    // Generates a JSON file with a relatively large array of person objects
    private static void CreateLargeJson(string filePath)
    {
        using (StreamWriter writer = new StreamWriter(filePath))
        {
            writer.Write("[");
            for (int i = 1; i <= 1000; i++)
            {
                writer.Write($"{{\"Name\":\"Person {i}\",\"Age\":{20 + i % 30}}}");
                if (i < 1000) writer.Write(",");
            }
            writer.Write("]");
        }
    }

    // Generates a JSON file with a small array of person objects
    private static void CreateSmallJson(string filePath)
    {
        using (StreamWriter writer = new StreamWriter(filePath))
        {
            writer.Write("[");
            for (int i = 1; i <= 5; i++)
            {
                writer.Write($"{{\"Name\":\"Mini Person {i}\",\"Age\":{25 + i}}}");
                if (i < 5) writer.Write(",");
            }
            writer.Write("]");
        }
    }
}
