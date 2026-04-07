using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportData
{
    // Flag that determines whether the appendix should be merged.
    public bool IncludeAppendix { get; set; }

    // The document to be appended when the flag is true.
    public Document AppendixDoc { get; set; } = null!;

    // Timestamp used in the template.
    public string GeneratedOn { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Paths for temporary files.
        const string templatePath = "Template.docx";
        const string appendixPath = "Appendix.docx";
        const string resultPath = "Result.docx";

        // -------------------------------------------------
        // 1. Create the appendix document.
        // -------------------------------------------------
        Document appendixDoc = new Document();
        DocumentBuilder appendixBuilder = new DocumentBuilder(appendixDoc);
        appendixBuilder.Writeln("=== Appendix ===");
        appendixBuilder.Writeln("This content is added only when the flag is true.");
        // Save the appendix so it can be loaded later.
        appendixDoc.Save(appendixPath);

        // -------------------------------------------------
        // 2. Create the main template document with LINQ Reporting tags.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);

        templateBuilder.Writeln("Report Title");
        templateBuilder.Writeln("Generated on: <<[model.GeneratedOn]>>");
        templateBuilder.Writeln(); // Empty line for readability.

        // Conditional block: include appendix only when IncludeAppendix is true.
        templateBuilder.Writeln("<<if [model.IncludeAppendix]>>");
        // The <<doc>> tag inserts another document.
        templateBuilder.Writeln("<<doc [model.AppendixDoc]>>");
        templateBuilder.Writeln("<</if>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 3. Load the documents back (simulating a real scenario).
        // -------------------------------------------------
        Document loadedTemplate = new Document(templatePath);
        Document loadedAppendix = new Document(appendixPath);

        // -------------------------------------------------
        // 4. Prepare the data model.
        // -------------------------------------------------
        var data = new ReportData
        {
            IncludeAppendix = true,               // Change to false to skip the appendix.
            AppendixDoc = loadedAppendix,
            GeneratedOn = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        };

        // -------------------------------------------------
        // 5. Build the report using the LINQ Reporting engine.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // Use property assignment, not object initializer.

        // The root name in the template is "model".
        engine.BuildReport(loadedTemplate, data, "model");

        // -------------------------------------------------
        // 6. Save the final document.
        // -------------------------------------------------
        loadedTemplate.Save(resultPath);
    }
}
