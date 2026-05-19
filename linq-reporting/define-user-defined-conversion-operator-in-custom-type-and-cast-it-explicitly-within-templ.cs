using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class MyNumber
{
    public int Value { get; set; }

    public MyNumber(int value) => Value = value;

    // Explicit conversion from MyNumber to int
    public static explicit operator int(MyNumber number) => number.Value;
}

public class ReportModel
{
    // Initialize the collection with sample data to avoid nullable warnings
    public List<MyNumber> Items { get; set; } = new()
    {
        new MyNumber(10),
        new MyNumber(20),
        new MyNumber(30)
    };
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create the template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        const string templatePath = "Template.docx";

        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin a foreach loop over the Items collection of the root object "model".
        builder.Writeln("<<foreach [item in model.Items]>>");
        // Explicitly cast the custom type to int within the expression.
        builder.Writeln("Value: <<[(int)item]>>");
        builder.Writeln("<</foreach>>");

        // Save the template so that it can be loaded for the report generation step.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Prepare the data source.
        ReportModel model = new();

        // Configure and execute the reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // No special options needed.
        bool success = engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);

        // Optional: indicate success (no console interaction required).
        // The program will exit automatically.
    }
}
