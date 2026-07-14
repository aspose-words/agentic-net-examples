using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingRestrictedMembers
{
    // Sample data model with a method that modifies its state.
    public class SampleModel
    {
        // Initialize to avoid nullable warnings.
        public int Counter { get; set; } = 0;

        // Method that changes the Counter property.
        public int Increment()
        {
            Counter++;
            return Counter;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare a simple template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Output the current counter value.
            builder.Writeln("Counter value: <<[model.Counter]>>");
            // Attempt to call a method that modifies the data source.
            builder.Writeln("Attempt to increment: <<[model.Increment()]>>");

            // Save the template to a local file.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Load the template back (simulating a separate load step).
            Document doc = new Document(templatePath);

            // Create the data source instance.
            SampleModel model = new SampleModel();

            // Restrict access to the SampleModel type so its members cannot be used in the template.
            // This must be done before any report is built.
            ReportingEngine.SetRestrictedTypes(typeof(SampleModel));

            // Configure the reporting engine to inline error messages.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.InlineErrorMessages
            };

            // Build the report. The method call will be blocked and an error message will appear inline.
            engine.BuildReport(doc, model, "model");

            // Save the resulting document.
            const string outputPath = "Report_Output.docx";
            doc.Save(outputPath);
        }
    }
}
