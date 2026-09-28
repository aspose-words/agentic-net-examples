using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public bool IncludeAppendix { get; set; }
    public string AppendixPath { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // File paths
        const string templatePath = "Template.docx";
        const string appendixPath = "Appendix.docx";
        const string outputPath = "Result.docx";

        // Create the main template document with a conditional include for the appendix.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Main Report Content");
        builder.Writeln("<<if [IncludeAppendix]>>");
        builder.Writeln("<<doc [AppendixPath]>>");
        builder.Writeln("<</if>>");
        templateDoc.Save(templatePath);

        // Create a simple appendix document.
        var appendixDoc = new Document();
        var appendixBuilder = new DocumentBuilder(appendixDoc);
        appendixBuilder.Writeln("Appendix Content");
        appendixDoc.Save(appendixPath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Prepare the data model.
        var model = new ReportModel
        {
            IncludeAppendix = true,          // Set to true to include the appendix.
            AppendixPath = appendixPath      // Path to the appendix document.
        };

        // Build the report using LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the final document.
        doc.Save(outputPath);
    }
}
