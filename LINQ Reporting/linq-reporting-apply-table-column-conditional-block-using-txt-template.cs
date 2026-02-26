using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model that will be used as the data source for the report.
    public class Product
    {
        public string Name { get; set; }
        public decimal Price { get; set; }
        public bool InStock { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Load the TXT template that contains a table with a conditional block.
            //    The template might look like this (saved as Template.txt):
            //    
            //    <<foreach [ds.Products]>>
            //    | <<[Name]>> | <<[Price]:currency>> | <<if [InStock]>><<[InStock]>><<else>>Out of stock<<endif>>
            //    <<endforeach>>
            //    
            //    Aspose.Words can load a plain‑text file directly into a Document object.
            // -----------------------------------------------------------------
            Document template = new Document("Template.txt");

            // -----------------------------------------------------------------
            // 2. Prepare the data source.
            //    The ReportingEngine can work with any non‑dynamic, non‑anonymous .NET type.
            // -----------------------------------------------------------------
            var data = new
            {
                Products = new List<Product>
                {
                    new Product { Name = "Apple",  Price = 0.99m, InStock = true  },
                    new Product { Name = "Banana", Price = 0.59m, InStock = false },
                    new Product { Name = "Cherry", Price = 2.49m, InStock = true  }
                }
            };

            // -----------------------------------------------------------------
            // 3. Create and configure the ReportingEngine.
            //    - RemoveEmptyParagraphs: cleans up any empty paragraphs that may appear
            //      after conditional blocks evaluate to false.
            //    - InlineErrorMessages: makes debugging easier by inserting error messages
            //      directly into the output document if the template contains syntax errors.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.RemoveEmptyParagraphs |
                          ReportBuildOptions.InlineErrorMessages
            };

            // -----------------------------------------------------------------
            // 4. Build the report.
            //    Use the overload that allows us to reference the data source object
            //    itself inside the template (the name "ds" is used in the template tags).
            // -----------------------------------------------------------------
            bool success = engine.BuildReport(template, data, "ds");

            // Optional: check the return value if InlineErrorMessages option is set.
            if (!success)
            {
                Console.WriteLine("The template contains errors. See the generated document for details.");
            }

            // -----------------------------------------------------------------
            // 5. Save the populated document.
            // -----------------------------------------------------------------
            template.Save("Report.docx");
        }
    }
}
