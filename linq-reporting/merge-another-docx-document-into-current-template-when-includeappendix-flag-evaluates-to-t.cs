using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        const string templatePath = "Template.docx";
        const string appendixPath = "Appendix.docx";
        const string resultPath = "Result.docx";

        // -----------------------------------------------------------------
        // 1. Create the appendix document that may be merged later.
        // -----------------------------------------------------------------
        Document appendixDoc = new Document();
        DocumentBuilder appendixBuilder = new DocumentBuilder(appendixDoc);
        appendixBuilder.Writeln("Appendix Content");
        appendixDoc.Save(appendixPath);

        // -----------------------------------------------------------------
        // 2. Create the main template with a conditional <<doc>> tag.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
        templateBuilder.Writeln("Main Report Content");
        // Conditional block: include appendix only when the flag is true.
        templateBuilder.Writeln("<<if [model.IncludeAppendix]>>");
        templateBuilder.Writeln("<<doc [model.Appendix]>>");
        templateBuilder.Writeln("<</if>>");
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template (as required by the lifecycle rules).
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 4. Prepare the data model.
        // -----------------------------------------------------------------
        ReportModel model = new()
        {
            IncludeAppendix = true,               // Flag that controls inclusion.
            Appendix = new Document(appendixPath) // Document to be merged.
        };

        // -----------------------------------------------------------------
        // 5. Build the report using Aspose.Words LINQ Reporting Engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 6. Save the final document.
        // -----------------------------------------------------------------
        reportDoc.Save(resultPath);
    }
}

// Data model aligned with the template tags.
public class ReportModel
{
    // Determines whether the appendix should be inserted.
    public bool IncludeAppendix { get; set; } = false;

    // The document to be merged when IncludeAppendix is true.
    public Document Appendix { get; set; } = null!;
}
