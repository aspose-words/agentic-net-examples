using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class MyNumber
{
    public int Value { get; set; }

    public MyNumber(int value) => Value = value;

    // Explicit conversion to int
    public static explicit operator int(MyNumber number) => number.Value;
}

public class ReportModel
{
    // Initialize to avoid nullable warnings
    public MyNumber MyNumber { get; set; } = new MyNumber(0);
    public string Title { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data
        var model = new ReportModel
        {
            Title = "Conversion Demo",
            MyNumber = new MyNumber(42)
        };

        // Create template document programmatically
        var templatePath = "Template.docx";
        var builder = new DocumentBuilder();
        builder.Writeln("<<[model.Title]>>");
        // Explicit cast of custom type to int inside the template expression
        builder.Writeln("Number (cast to int): <<[(int)model.MyNumber]>>");
        builder.Document.Save(templatePath);

        // Load the template for reporting
        var doc = new Document(templatePath);

        // Build the report using LINQ Reporting engine
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        var outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
