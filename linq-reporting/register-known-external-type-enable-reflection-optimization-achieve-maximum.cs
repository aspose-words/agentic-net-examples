using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReflectionOptimizationDemo
{
    // Sample external type whose static members can be accessed from a report template.
    public static class ExternalHelper
    {
        public static string GetGreeting(string name) => $"Hello, {name}!";
    }

    class Program
    {
        static void Main()
        {
            // Enable reflection optimization for maximum performance.
            // This causes the engine to generate dynamic classes for faster member access.
            ReportingEngine.UseReflectionOptimization = true;

            // Create a new reporting engine instance.
            ReportingEngine engine = new ReportingEngine();

            // Register the external type so that its static members can be used in templates.
            engine.KnownTypes.Add(typeof(ExternalHelper));

            // Prepare a simple template document.
            // The template uses the registered external type to call GetGreeting().
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);
            builder.Writeln("<<= ExternalHelper.GetGreeting(FirstName) >>");

            // Sample data source – a list of anonymous objects.
            var data = new List<object>
            {
                new { FirstName = "Alice" },
                new { FirstName = "Bob" },
                new { FirstName = "Charlie" }
            };

            // Build the report by populating the template with the data source.
            engine.BuildReport(template, data.ToArray());

            // Save the resulting document.
            template.Save("ReportWithReflectionOptimization.docx");
        }
    }
}
