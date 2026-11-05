using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

public class Program
{
    public static void Main()
    {
        // Register code page provider required by Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a template document containing a LINQ Reporting tag.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Customer: <<[model.Name]>>");

        // Save and reload the template to ensure it is fully created before building the report.
        const string templatePath = "Template.docx";
        template.Save(templatePath);
        Document doc = new Document(templatePath);

        // Prepare the data model.
        ReportModel model = new ReportModel { Name = "John Doe" };

        // Disable reflection optimization for this scoped operation.
        bool previousSetting = ReportingEngine.UseReflectionOptimization;
        ReportingEngine.UseReflectionOptimization = false;
        try
        {
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");
        }
        finally
        {
            // Restore the original setting.
            ReportingEngine.UseReflectionOptimization = previousSetting;
        }

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }

    // Public data model class.
    public class ReportModel
    {
        public string Name { get; set; } = "";
    }
}
