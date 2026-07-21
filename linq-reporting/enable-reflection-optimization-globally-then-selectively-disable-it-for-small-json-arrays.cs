using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting; // Contains ReportingEngine, JsonDataSource, etc.

public class ReflectionOptimizationDemo
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a template document with two data sections: Large and Small.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(outputDir, "Template.docx");
        CreateTemplate(templatePath);

        // -----------------------------------------------------------------
        // 2. Create a JSON file containing a large array and a small array.
        // -----------------------------------------------------------------
        string jsonPath = Path.Combine(outputDir, "Data.json");
        CreateJsonData(jsonPath);

        // -----------------------------------------------------------------
        // 3. Global reflection optimization enabled.
        // -----------------------------------------------------------------
        ReportingEngine.UseReflectionOptimization = true; // enable globally

        Document docGlobal = new Document(templatePath);
        JsonDataSource dataSource = new JsonDataSource(jsonPath);
        ReportingEngine engineGlobal = new ReportingEngine();
        engineGlobal.BuildReport(docGlobal, dataSource, "data"); // root name "data"
        string globalResult = Path.Combine(outputDir, "Result_GlobalOptimization.docx");
        docGlobal.Save(globalResult);

        // -----------------------------------------------------------------
        // 4. Disable reflection optimization for small JSON arrays.
        //    (Demonstrated by building a report that only uses the Small array.)
        // -----------------------------------------------------------------
        ReportingEngine.UseReflectionOptimization = false; // disable for this scenario

        // Create a template that only references the Small array.
        string smallTemplatePath = Path.Combine(outputDir, "Template_SmallOnly.docx");
        CreateSmallOnlyTemplate(smallTemplatePath);

        Document docSmall = new Document(smallTemplatePath);
        JsonDataSource smallDataSource = new JsonDataSource(jsonPath);
        ReportingEngine engineSmall = new ReportingEngine();
        engineSmall.BuildReport(docSmall, smallDataSource, "data");
        string smallResult = Path.Combine(outputDir, "Result_SmallArray_NoOptimization.docx");
        docSmall.Save(smallResult);

        // Reset the static flag to its default (true) for any further operations.
        ReportingEngine.UseReflectionOptimization = true;
    }

    // Creates a template with tags for both Large and Small arrays.
    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Large items:");
        builder.Writeln("<<foreach [item in Large]>>");
        builder.Writeln("- <<[item]>>");
        builder.Writeln("<</foreach>>");

        builder.Writeln();
        builder.Writeln("Small items:");
        builder.Writeln("<<foreach [item in Small]>>");
        builder.Writeln("- <<[item]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(filePath);
    }

    // Creates a template that only references the Small array.
    private static void CreateSmallOnlyTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Small items (no reflection optimization):");
        builder.Writeln("<<foreach [item in Small]>>");
        builder.Writeln("- <<[item]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(filePath);
    }

    // Generates a JSON file with a large array (100 items) and a small array (2 items).
    private static void CreateJsonData(string filePath)
    {
        var jsonObject = new
        {
            Large = GenerateNumbers(1, 100),
            Small = GenerateNumbers(1, 2)
        };

        string json = System.Text.Json.JsonSerializer.Serialize(
            jsonObject,
            new System.Text.Json.JsonSerializerOptions { WriteIndented = true });

        File.WriteAllText(filePath, json);
    }

    // Helper to generate a list of integers as strings.
    private static List<string> GenerateNumbers(int start, int count)
    {
        var list = new List<string>();
        for (int i = start; i < start + count; i++)
            list.Add(i.ToString());
        return list;
    }
}
