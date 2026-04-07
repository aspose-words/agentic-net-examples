using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    public class Program
    {
        public static void Main()
        {
            // Create a template document with LINQ Reporting tags.
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Insert a greeting that uses a data field.
            builder.Writeln("Report for <<[model.Name]>>");
            builder.Writeln();

            // Section visible only when the role is Admin.
            builder.Writeln("<<if [model.Role == \"Admin\"]>>");
            builder.Writeln("=== Admin Section ===");
            builder.Writeln("Confidential data for administrators.");
            builder.Writeln("<</if>>");

            // Section visible for all other roles.
            builder.Writeln("<<if [model.Role != \"Admin\"]>>");
            builder.Writeln("=== User Section ===");
            builder.Writeln("General data for regular users.");
            builder.Writeln("<</if>>");

            // Save the template to disk (required before building the report).
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Load the template document.
            var doc = new Document(templatePath);

            // Prepare the data source.
            var model = new ReportModel
            {
                Name = "John Doe",
                Role = "Admin"
            };

            // Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save("Report.docx");
        }
    }

    // Simple data model used by the template.
    public class ReportModel
    {
        public string Name { get; set; } = "";
        public string Role { get; set; } = "";
    }
}
