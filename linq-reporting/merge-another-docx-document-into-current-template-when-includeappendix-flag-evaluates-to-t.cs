using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportData
{
    // Flag that determines whether the appendix should be merged.
    public bool IncludeAppendix { get; set; } = true;

    // The appendix document to be inserted when the flag is true.
    public Document Appendix { get; set; } = null!;
}

// Public model class used by the LINQ Reporting engine.
// Must be a concrete (non‑anonymous) type for BuildReport.
public class ReportModel
{
    public bool IncludeAppendix { get; set; }
    public Document Appendix { get; set; } = null!;
    public string GenerationDate { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // ---------------------------------------------------------------
        // 1. Create the appendix document that will be merged conditionally.
        // ---------------------------------------------------------------
        Document appendixDoc = new Document();
        var appendixBuilder = new DocumentBuilder(appendixDoc);
        appendixBuilder.Writeln("=== Appendix ===");
        appendixBuilder.Writeln("This is the appended content.");
        // Save the appendix for reference (optional).
        string appendixPath = Path.Combine(outputDir, "Appendix.docx");
        appendixDoc.Save(appendixPath);

        // ---------------------------------------------------------------
        // 2. Create the main template document containing LINQ Reporting tags.
        // ---------------------------------------------------------------
        Document templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report Title");
        builder.Writeln("Generated on: <<[model.GenerationDate]>>");
        builder.Writeln(string.Empty); // Empty line for readability.

        // Conditional block: include the appendix only when IncludeAppendix is true.
        builder.Writeln("<<if [model.IncludeAppendix]>>");
        builder.Writeln("<<doc [model.Appendix]>>");
        builder.Writeln("<</if>>");

        // Save the template (optional, useful for debugging).
        string templatePath = Path.Combine(outputDir, "Template.docx");
        templateDoc.Save(templatePath);

        // ---------------------------------------------------------------
        // 3. Prepare the data model for the report.
        // ---------------------------------------------------------------
        var model = new ReportModel
        {
            IncludeAppendix = true,               // Change to false to skip the appendix.
            Appendix = appendixDoc,                // Document to be merged.
            GenerationDate = DateTime.Now.ToString("yyyy-MM-dd")
        };

        // ---------------------------------------------------------------
        // 4. Build the final report using the LINQ Reporting engine.
        // ---------------------------------------------------------------
        var engine = new ReportingEngine();
        // The template uses the root name "model", so we pass the model accordingly.
        engine.BuildReport(templateDoc, model, "model");

        // ---------------------------------------------------------------
        // 5. Save the generated report.
        // ---------------------------------------------------------------
        string resultPath = Path.Combine(outputDir, "FinalReport.docx");
        templateDoc.Save(resultPath);
    }
}
