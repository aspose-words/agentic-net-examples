using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Flag that determines whether the appendix should be included.
    public bool IncludeAppendix { get; set; } = true;

    // The appendix document to be merged when the flag is true.
    public Document AppendixDocument { get; set; } = null!;
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create the main template document with a conditional <<doc>> tag.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder tmplBuilder = new DocumentBuilder(template);

        tmplBuilder.Writeln("=== Main Report ===");
        tmplBuilder.Writeln("This is the main part of the report.");

        // Conditional block: include the appendix only when IncludeAppendix is true.
        tmplBuilder.Writeln("<<if [model.IncludeAppendix]>>");
        tmplBuilder.Writeln("<<doc [model.AppendixDocument]>>");
        tmplBuilder.Writeln("<</if>>");

        // Save the template to disk (required by the workflow).
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Create the appendix document that will be merged.
        // -----------------------------------------------------------------
        Document appendix = new Document();
        DocumentBuilder appBuilder = new DocumentBuilder(appendix);
        appBuilder.Writeln("=== Appendix ===");
        appBuilder.Writeln("This content is conditionally appended to the main report.");
        const string appendixPath = "Appendix.docx";
        appendix.Save(appendixPath);

        // -----------------------------------------------------------------
        // 3. Load the documents back (simulating a real-world scenario).
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);
        Document loadedAppendix = new Document(appendixPath);

        // -----------------------------------------------------------------
        // 4. Prepare the data model.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            IncludeAppendix = true,               // Change to false to omit the appendix.
            AppendixDocument = loadedAppendix
        };

        // -----------------------------------------------------------------
        // 5. Build the report using the LINQ Reporting engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // -----------------------------------------------------------------
        // 6. Save the final merged document.
        // -----------------------------------------------------------------
        const string resultPath = "Result.docx";
        loadedTemplate.Save(resultPath);
    }
}
