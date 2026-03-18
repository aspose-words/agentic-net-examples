using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Preserve the original setting so it can be restored later.
            bool originalOptimization = ReportingEngine.UseReflectionOptimization;

            try
            {
                // Disable reflection optimization for the duration of this block.
                ReportingEngine.UseReflectionOptimization = false;

                // Load the template document.
                Document template = new Document("Template.docx");

                // Prepare the data source for the report.
                var data = new
                {
                    // Example data – replace with your actual model.
                    Title = "Sample Report",
                    Items = new[]
                    {
                        new { Name = "Item 1", Quantity = 10 },
                        new { Name = "Item 2", Quantity = 20 }
                    }
                };

                // Build the report using the LINQ Reporting Engine.
                ReportingEngine engine = new ReportingEngine();
                engine.BuildReport(template, data);

                // Save the generated report.
                template.Save("Result.docx");
            }
            finally
            {
                // Restore the original optimization setting.
                ReportingEngine.UseReflectionOptimization = originalOptimization;
            }
        }
    }
}
