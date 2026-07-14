using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Both flags are false, so no conditional block will be true.
    public bool ShowA { get; set; } = false;
    public bool ShowB { get; set; } = false;
}

public class Program
{
    public static void Main()
    {
        // Create a blank document that will serve as the template.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Write a paragraph that contains two conditional blocks.
        // If neither condition is true, the paragraph becomes empty.
        builder.Writeln("<<if [model.ShowA]>>A<</if>> <<if [model.ShowB]>>B<</if>>");

        // Prepare the data source.
        ReportModel model = new ReportModel();

        // Configure the reporting engine to remove empty paragraphs after processing.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Build the report. The root object name must match the tag prefix ("model").
        engine.BuildReport(template, model, "model");

        // Save the result. The output file will not contain the empty paragraph.
        template.Save("ConditionalBlockRemoved.docx");
    }
}
