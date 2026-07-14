using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Flag to control inclusion of the appendix.
    public bool IncludeAppendix { get; set; } = false;

    // The document that will be inserted when the flag is true.
    public Document Appendix { get; set; } = new();

    // Timestamp that will be displayed in the template.
    public string GeneratedOn { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // File names for the temporary documents.
        const string templatePath = "Template.docx";
        const string appendixPath = "Appendix.docx";
        const string resultPath   = "Result.docx";

        // ---------- Create the appendix document ----------
        var appendixDoc = new Document();
        var appendixBuilder = new DocumentBuilder(appendixDoc);
        appendixBuilder.Writeln("=== Appendix ===");
        appendixBuilder.Writeln("This is the appended appendix content.");
        appendixDoc.Save(appendixPath);

        // ---------- Create the main template document ----------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("=== Main Report ===");
        builder.Writeln("Report generated on: <<[model.GeneratedOn]>>");
        // Conditional inclusion of the appendix.
        builder.Writeln("<<if [model.IncludeAppendix]>>");
        builder.Writeln("<<doc [model.Appendix]>>");
        builder.Writeln("<</if>>");
        templateDoc.Save(templatePath);

        // ---------- Load documents ----------
        var loadedTemplate = new Document(templatePath);
        var loadedAppendix = new Document(appendixPath);

        // ---------- Prepare data model ----------
        var model = new ReportModel
        {
            IncludeAppendix = true, // Set to false to omit the appendix.
            Appendix = loadedAppendix,
            GeneratedOn = DateTime.Now.ToString("yyyy-MM-dd HH:mm")
        };

        // ---------- Build the report ----------
        var engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // ---------- Save the final document ----------
        loadedTemplate.Save(resultPath);
    }
}
