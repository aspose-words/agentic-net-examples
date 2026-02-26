using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsReportingMhtml
{
    // Simple data source class.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a template document with LINQ Reporting tags.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);
            // Use contextual object member access to refer to the data source object named "person".
            builder.Writeln("<<[person.Name]>> is <<[person.Age]>> years old.");

            // 2. Prepare the data source.
            Person person = new Person { Name = "John Doe", Age = 30 };

            // 3. Build the report using the overload that allows referencing the data source object.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, person, "person");

            // 4. Save the populated document as MHTML.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
            {
                // Export resources (images, CSS, etc.) as CID URLs suitable for MHTML.
                ExportCidUrlsForMhtmlResources = true
            };
            template.Save("Report.mhtml", saveOptions);
        }
    }
}
