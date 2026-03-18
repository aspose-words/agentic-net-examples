using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace ReportingEngineExample
{
    // Simple data source class – some members may be missing in the template.
    public class SampleData
    {
        public string Name { get; set; }
        public int Age { get; set; }
        // Note: No property called "Address" – this will trigger the missing‑member placeholder.
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a template document in memory.
            // The template contains a plain reference to a missing member <<[data.Address]>>.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);
            builder.Writeln("Name: <<[data.Name]>>");
            builder.Writeln("Age: <<[data.Age]>>");
            builder.Writeln("Address: <<[data.Address]>>"); // Address does not exist in SampleData.

            // 2. Configure the ReportingEngine.
            ReportingEngine engine = new ReportingEngine
            {
                // Enable handling of missing members.
                Options = ReportBuildOptions.AllowMissingMembers,
                // Insert "<<error>>" wherever a member is missing.
                MissingMemberMessage = "<<error>>"
            };

            // 3. Build the report using a data source instance.
            SampleData data = new SampleData { Name = "John Doe", Age = 30 };
            // The data source name "data" matches the name used in the template tags.
            engine.BuildReport(template, data, "data");

            // 4. Save the resulting document.
            template.Save("ReportWithMissingMemberPlaceholders.docx");
        }
    }
}
