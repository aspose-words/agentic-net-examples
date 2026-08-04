using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for proper Unicode handling.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for the JSON source, template, and final report.
        const string jsonPath = "report.json";
        const string templatePath = "template.docx";
        const string outputPath = "report_output.docx";

        // -----------------------------------------------------------------
        // 1. Create a sample JSON file containing multilingual text.
        // -----------------------------------------------------------------
        var jsonContent = new
        {
            Title = "Multilingual Report",
            Items = new[]
            {
                new
                {
                    Name = "Apple",
                    Description_en = "Fresh apple",
                    Description_es = "Manzana fresca",
                    Description_zh = "新鲜的苹果",
                    Description_ar = "تفاحة طازجة"
                },
                new
                {
                    Name = "Banana",
                    Description_en = "Ripe banana",
                    Description_es = "Plátano maduro",
                    Description_zh = "成熟的香蕉",
                    Description_ar = "موز ناضج"
                }
            }
        };
        // Serialize the object to JSON and write it to a file.
        string jsonString = System.Text.Json.JsonSerializer.Serialize(
            jsonContent,
            new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(jsonPath, jsonString, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Build a Word template programmatically with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Report title.
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln();

        // Begin a foreach loop over the Items collection.
        builder.Writeln("<<foreach [item in model.Items]>>");
        builder.Writeln("Name: <<[item.Name]>>");
        builder.Writeln("English: <<[item.Description_en]>>");
        builder.Writeln("Spanish: <<[item.Description_es]>>");
        builder.Writeln("Chinese: <<[item.Description_zh]>>");
        builder.Writeln("Arabic: <<[item.Description_ar]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and generate the report using the JSON data source.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        JsonDataSource dataSource = new JsonDataSource(jsonPath);

        ReportingEngine engine = new ReportingEngine();
        // The root object name used in the template tags is "model".
        engine.BuildReport(reportDoc, dataSource, "model");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(outputPath);
    }
}
