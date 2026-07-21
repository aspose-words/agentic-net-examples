using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a new blank document that will serve as the template.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a LINQ Reporting tag that references a member which does NOT exist in the data source.
        // With AllowMissingMembers enabled, this will be treated as a null literal.
        builder.Writeln("Customer name: <<[Missing.Name]>>");

        // Save the template to disk so that it can be loaded later for report generation.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template back into a Document object.
        Document doc = new Document(templatePath);

        // Prepare a data source that does NOT contain the "Missing" object.
        // An empty DataSet is sufficient for this demonstration.
        DataSet data = new DataSet();

        // Configure the ReportingEngine to allow missing members.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;
        // Optional: customize the message that will be inserted for missing members.
        engine.MissingMemberMessage = "N/A";

        // Build the report. The missing field will be treated as null (empty) because of the option set above.
        engine.BuildReport(doc, data, "");

        // Save the generated report.
        const string reportPath = "Report.docx";
        doc.Save(reportPath);
    }
}
