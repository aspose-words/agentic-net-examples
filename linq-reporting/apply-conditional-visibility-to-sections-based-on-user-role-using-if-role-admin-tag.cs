using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Data model used by the LINQ Reporting template.
    public class ReportModel
    {
        // Role of the user (e.g., "Admin" or "User").
        public string Role { get; set; } = string.Empty;

        // Title displayed at the top of the report.
        public string Title { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a Word document that will serve as the template.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a title placeholder.
            builder.Writeln("<<[model.Title]>>");
            builder.Writeln();

            // Section visible only to administrators.
            builder.Writeln("<<if [model.Role == \"Admin\"]>>");
            builder.Writeln("=== Admin Section ===");
            builder.Writeln("Confidential data visible only to administrators.");
            builder.Writeln("<</if>>");
            builder.Writeln();

            // Section visible to non‑administrators.
            builder.Writeln("<<if [model.Role != \"Admin\"]>>");
            builder.Writeln("=== User Section ===");
            builder.Writeln("General information visible to all users.");
            builder.Writeln("<</if>>");

            // -----------------------------------------------------------------
            // 2. Prepare the data model.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                Role = "Admin",               // Change to "User" to see the other section.
                Title = "Monthly Report"
            };

            // -----------------------------------------------------------------
            // 3. Build the report using the LINQ Reporting engine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            template.Save("GeneratedReport.docx");
        }
    }
}
