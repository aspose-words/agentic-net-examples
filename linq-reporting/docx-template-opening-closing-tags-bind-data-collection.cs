using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsTemplateExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to insert the opening and closing tags that define a repeating region.
            // The syntax <<[CollectionName]>> starts the region, and <<[/CollectionName]>> ends it.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Opening tag for the collection named "Items".
            builder.Writeln("<<[Items]>>");

            // Content that will be repeated for each item in the collection.
            // Inside the region you can reference members of each item, e.g., Name and Quantity.
            builder.Writeln("Item: <<[Name]>>");
            builder.Writeln("Quantity: <<[Quantity]>>");
            builder.Writeln(); // Add an empty line between items.

            // Closing tag for the collection.
            builder.Writeln("<<[/Items]>>");

            // Save the template document.
            doc.Save("TemplateWithCollectionTags.docx");

            // Example of populating the template using ReportingEngine.
            // Define a simple data source class.
            var data = new
            {
                Items = new[]
                {
                    new { Name = "Apple", Quantity = 5 },
                    new { Name = "Banana", Quantity = 12 },
                    new { Name = "Cherry", Quantity = 20 }
                }
            };

            // Load the template (could also reuse the 'doc' instance).
            Document template = new Document("TemplateWithCollectionTags.docx");

            // Build the report by merging the data source with the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, data, ""); // Empty name because we reference members directly.

            // Save the populated report.
            template.Save("ReportGenerated.docx");
        }
    }
}
