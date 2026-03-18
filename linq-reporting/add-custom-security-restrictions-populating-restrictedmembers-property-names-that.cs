using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace ReportingSecurityDemo
{
    // Sample data class whose members we want to restrict.
    public class SensitiveData
    {
        public string Name { get; set; }
        public string SSN { get; set; }               // Sensitive
        public decimal Salary { get; set; }            // Sensitive
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a simple template document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Name: <<[data.Name]>>");
            builder.Writeln("SSN: <<[data.SSN]>>");          // This member will be hidden.
            builder.Writeln("Salary: <<[data.Salary]>>");    // This member will be hidden.

            // 2. Define the data source.
            SensitiveData data = new SensitiveData
            {
                Name = "John Doe",
                SSN = "123-45-6789",
                Salary = 85000m
            };

            // 3. Restrict the SensitiveData type so its members cannot be accessed
            //    from the reporting engine. This effectively hides the SSN and Salary
            //    properties in the generated report.
            ReportingEngine.SetRestrictedTypes(typeof(SensitiveData));

            // 4. Build the report. Use AllowMissingMembers to avoid exceptions for
            //    the restricted members – they will be rendered as empty strings.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.AllowMissingMembers
            };
            engine.BuildReport(doc, data, "data");

            // 5. Save the result.
            doc.Save("ReportWithRestrictedMembers.docx");

            // The resulting document will contain:
            // Name: John Doe
            // SSN:
            // Salary:
        }
    }
}
