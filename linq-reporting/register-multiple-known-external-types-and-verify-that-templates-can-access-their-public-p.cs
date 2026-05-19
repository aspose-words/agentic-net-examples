using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // External type 1 with a public static property.
    public static class ExternalTypeA
    {
        public static string Message => "Hello from ExternalTypeA";
    }

    // External type 2 with a public static property.
    public static class ExternalTypeB
    {
        public static int Value => 12345;
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document that contains tags referencing the
            //    static properties of the external types.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Message: <<[ExternalTypeA.Message]>>");
            builder.Writeln("Value:   <<[ExternalTypeB.Value]>>");

            // Save the template to disk.
            const string templatePath = "template.docx";
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back (simulating a real‑world scenario where the
            //    template is stored separately).
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Configure the ReportingEngine.
            //    Register the external types so that their static members can be
            //    accessed from the template without using reflection.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.KnownTypes.Add(typeof(ExternalTypeA));
            engine.KnownTypes.Add(typeof(ExternalTypeB));

            // The data source is not used in this example because the template only
            // accesses static members. An empty object is sufficient.
            object dummyDataSource = new object();

            // Build the report.
            engine.BuildReport(loadedTemplate, dummyDataSource, "data");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            const string resultPath = "result.docx";
            loadedTemplate.Save(resultPath);

            // -----------------------------------------------------------------
            // 5. Verify the output by printing the document text to the console.
            // -----------------------------------------------------------------
            Console.WriteLine("Generated report content:");
            Console.WriteLine(loadedTemplate.GetText());
        }
    }
}
