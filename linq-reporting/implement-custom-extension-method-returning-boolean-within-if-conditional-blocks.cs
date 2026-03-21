using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Fields;

namespace MyTemplateExtensions
{
    // Custom static class containing an extension method usable in IF field conditions.
    public static class TemplateExtensions
    {
        // Returns true if the supplied integer is even, false otherwise.
        public static bool IsEven(int number)
        {
            return number % 2 == 0;
        }
    }

    // Simple data source class with a public property.
    public class DataSource
    {
        public int Count { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a document with an IF field that uses the custom method.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("{ IF TemplateExtensions.IsEven(Count) \"Even\" \"Odd\" }");

            // Prepare the data source.
            var data = new DataSource { Count = 5 };

            // Configure the reporting engine to recognize the static class containing the custom method.
            ReportingEngine engine = new ReportingEngine();
            engine.KnownTypes.Add(typeof(TemplateExtensions));

            // Build the report – the placeholder TemplateExtensions.IsEven(Count) will be evaluated,
            // its boolean result will be inserted into the IF field, which then displays the appropriate text.
            engine.BuildReport(doc, data);

            // Save the populated document.
            doc.Save("Result.docx");
        }
    }
}
