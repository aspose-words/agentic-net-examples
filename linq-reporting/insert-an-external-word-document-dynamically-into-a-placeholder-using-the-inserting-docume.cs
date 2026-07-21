using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // The external document to be inserted.
    public Document ExternalDoc { get; set; } = null!;
}

public class Program
{
    public static void Main()
    {
        // Paths for the files used in the example.
        const string externalPath = "External.docx";
        const string templatePath = "Template.docx";
        const string resultPath = "Result.docx";

        // -----------------------------------------------------------------
        // 1. Create the external document that will be inserted later.
        // -----------------------------------------------------------------
        Document externalDoc = new Document();
        DocumentBuilder externalBuilder = new DocumentBuilder(externalDoc);
        externalBuilder.Writeln("This is the content of the external document.");
        externalDoc.Save(externalPath);

        // -----------------------------------------------------------------
        // 2. Create the template document containing a placeholder tag.
        //    The tag <<doc [model.ExternalDoc]>> tells the reporting engine
        //    to insert the document referenced by the ExternalDoc property.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
        templateBuilder.Writeln("Report start");
        templateBuilder.Writeln("<<doc [model.ExternalDoc]>>");
        templateBuilder.Writeln("Report end");
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template for reporting.
        // -----------------------------------------------------------------
        Document template = new Document(templatePath);

        // -----------------------------------------------------------------
        // 4. Load the external document that will be inserted.
        // -----------------------------------------------------------------
        Document external = new Document(externalPath);

        // -----------------------------------------------------------------
        // 5. Prepare the data model.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel { ExternalDoc = external };

        // -----------------------------------------------------------------
        // 6. Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // -----------------------------------------------------------------
        // 7. Save the final document.
        // -----------------------------------------------------------------
        template.Save(resultPath);
    }
}
