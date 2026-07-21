using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public bool Flag1 { get; set; } = false;
    public bool Flag2 { get; set; } = false;
    public bool Flag3 { get; set; } = false;
    public string Message { get; set; } = "No flags are true.";
}

public class Program
{
    public static void Main()
    {
        // Create a template document with conditional blocks.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Report Start");
        builder.Writeln("<<if [model.Flag1]>>Flag1 is true<</if>>");
        builder.Writeln("<<if [model.Flag2]>>Flag2 is true<</if>>");
        builder.Writeln("<<if [model.Flag3]>>Flag3 is true<</if>>");
        builder.Writeln("Report End");

        // Data source where all conditions are false.
        ReportModel model = new ReportModel();

        // Configure the engine to remove empty paragraphs after processing.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Build the report.
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("ConditionalReport.docx");
    }
}
