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

    public class Program
    {
        public static void Main()
        {
            // Load a template document that contains an IF field like:
            // { IF {{MyTemplateExtensions.TemplateExtensions.IsEven(Count)}} "Even" "Odd" }
            Document doc = new Document("Template.docx");

            // Data source for the template. The property name must match the placeholder used in the template.
            var data = new { Count = 5 };

            // Configure the reporting engine to recognize the static class containing the custom method.
            ReportingEngine engine = new ReportingEngine();
            engine.KnownTypes.Add(typeof(TemplateExtensions));

            // Build the report – the placeholder {{TemplateExtensions.IsEven(Count)}} will be evaluated,
            // its boolean result will be inserted into the IF field, which then displays the appropriate text.
            engine.BuildReport(doc, data);

            // Save the populated document.
            doc.Save("Result.docx");
        }
    }
}
