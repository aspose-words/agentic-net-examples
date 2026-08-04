using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingRestrictedMembers
{
    // Simple data model.
    public class Person
    {
        public string Name { get; set; } = "John Doe";
        public int Age { get; set; } = 30;
        public string Secret { get; set; } = "TopSecret";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "template.docx";
            const string outputPath = "report.docx";

            // -------------------------------------------------
            // 1. Create a template document with LINQ Reporting tags.
            // -------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Name: <<[Name]>>");
            builder.Writeln("Age: <<[Age]>>");
            builder.Writeln("Secret: <<[Secret]>>");
            // Attempt to access the System.Type of the object.
            // This will be blocked after we restrict System.Type.
            builder.Writeln("Type: <<[GetType().FullName]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template for reporting.
            // -------------------------------------------------
            var reportDoc = new Document(templatePath);

            // -------------------------------------------------
            // 3. Configure restricted members.
            //    Restrict System.Type so its members cannot be accessed from the template.
            // -------------------------------------------------
            ReportingEngine.SetRestrictedTypes(typeof(System.Type));

            // -------------------------------------------------
            // 4. Prepare the reporting engine.
            //    Allow missing members to avoid exceptions when a restricted member is accessed.
            // -------------------------------------------------
            var engine = new ReportingEngine
            {
                Options = ReportBuildOptions.AllowMissingMembers,
                MissingMemberMessage = "Restricted"
            };

            // -------------------------------------------------
            // 5. Create the data source.
            // -------------------------------------------------
            var person = new Person();

            // -------------------------------------------------
            // 6. Build the report.
            //    Use the overload without a data source name; the root object is the data source itself.
            // -------------------------------------------------
            engine.BuildReport(reportDoc, person);

            // -------------------------------------------------
            // 7. Save the generated report.
            // -------------------------------------------------
            reportDoc.Save(outputPath);
        }
    }
}
