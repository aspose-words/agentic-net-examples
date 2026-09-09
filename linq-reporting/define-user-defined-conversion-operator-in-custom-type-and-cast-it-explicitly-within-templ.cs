using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // 1. Create a template document with a LINQ Reporting tag that casts a custom type.
        string templatePath = Path.Combine(workDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Custom value: <<[(string)model.Custom]>>");
        templateDoc.Save(templatePath);

        // 2. Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // 3. Prepare the data model.
        ReportModel model = new ReportModel
        {
            Custom = new MyNumber(42) // The custom type will be cast to string in the template.
        };

        // 4. Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // 5. Save the generated report.
        string outputPath = Path.Combine(workDir, "Result.docx");
        reportDoc.Save(outputPath);
    }
}

// Public data model required by the template.
public class ReportModel
{
    // Initialized to avoid nullable warnings.
    public MyNumber Custom { get; set; } = null!;
}

// Custom type with an explicit conversion operator to string.
public class MyNumber
{
    public int Value { get; }

    public MyNumber(int value) => Value = value;

    // Explicit cast to string – used in the template expression.
    public static explicit operator string(MyNumber number) => number.Value.ToString();
}
