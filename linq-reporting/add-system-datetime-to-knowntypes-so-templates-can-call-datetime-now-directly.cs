using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Step 1: Create a template document with a LINQ Reporting tag that uses DateTime.Now.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Current date and time: <<[DateTime.Now]>>");
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Step 2: Load the template (simulating a separate load step).
        Document doc = new Document(templatePath);

        // Step 3: Configure the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // Add System.DateTime to the set of known types so the template can access static members.
        engine.KnownTypes.Add(typeof(DateTime));

        // Step 4: Build the report. No data source is needed because the template only uses a static member.
        // Pass a dummy object as the data source and an empty name.
        engine.BuildReport(doc, new object(), "");

        // Step 5: Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
