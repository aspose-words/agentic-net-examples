using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Holds the external document to be inserted.
    public Document External { get; set; } = new Document();
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create an external Word document that will be inserted later.
        // -----------------------------------------------------------------
        const string externalPath = "external.docx";
        Document externalDoc = new Document();
        DocumentBuilder extBuilder = new DocumentBuilder(externalDoc);
        extBuilder.Writeln("This is the content of the external document.");
        externalDoc.Save(externalPath);

        // -----------------------------------------------------------------
        // 2. Create the template document with a placeholder for insertion.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder tmplBuilder = new DocumentBuilder(template);
        tmplBuilder.Writeln("=== Report Start ===");
        // The placeholder uses the LINQ Reporting 'doc' tag.
        tmplBuilder.Writeln("<<doc [model.External]>>");
        tmplBuilder.Writeln("=== Report End ===");

        // -----------------------------------------------------------------
        // 3. Prepare the data model that supplies the external document.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            // Load the previously saved external document.
            External = new Document(externalPath)
        };

        // -----------------------------------------------------------------
        // 4. Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(template, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        const string outputPath = "output.docx";
        template.Save(outputPath);
    }
}
