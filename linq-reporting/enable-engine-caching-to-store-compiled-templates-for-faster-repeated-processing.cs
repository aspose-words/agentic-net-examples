using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for older encodings (required by Aspose.Words on .NET Core).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define file paths.
        const string templatePath = "Template.docx";
        const string outputPath1 = "Report1.docx";
        const string outputPath2 = "Report2.docx";

        // -----------------------------------------------------------------
        // 1. Create a simple template document with a LINQ Reporting tag.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Hello <<[model.Name]>>!"); // LINQ Reporting tag.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template for processing (two separate instances).
        // -----------------------------------------------------------------
        Document doc1 = new Document(templatePath);
        Document doc2 = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare sample data models.
        // -----------------------------------------------------------------
        var model1 = new ReportModel { Name = "Alice" };
        var model2 = new ReportModel { Name = "Bob" };

        // -----------------------------------------------------------------
        // 4. Create a ReportingEngine. Caching of compiled templates is
        //    performed automatically by the engine; no explicit Cache property
        //    is required in recent versions.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();

        // -----------------------------------------------------------------
        // 5. Build the first report (template will be compiled and cached).
        // -----------------------------------------------------------------
        engine.BuildReport(doc1, model1, "model");
        doc1.Save(outputPath1);

        // -----------------------------------------------------------------
        // 6. Build the second report using the same engine (cached template reused).
        // -----------------------------------------------------------------
        engine.BuildReport(doc2, model2, "model");
        doc2.Save(outputPath2);
    }

    // Simple data model referenced by the template.
    public class ReportModel
    {
        public string Name { get; set; } = string.Empty;
    }
}
