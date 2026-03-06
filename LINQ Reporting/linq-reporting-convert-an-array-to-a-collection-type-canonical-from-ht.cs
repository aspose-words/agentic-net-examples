using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading; // Needed for LoadOptions
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // HTML template that uses LINQ Reporting syntax to iterate over a collection named "items".
        string htmlTemplate = @"
        <html>
        <body>
        <<foreach [items]>><p>Item: <<[Current]>> </p><</foreach>>
        </body>
        </html>";

        // Load the HTML string into an Aspose.Words Document.
        Document templateDoc;
        using (MemoryStream htmlStream = new MemoryStream(Encoding.UTF8.GetBytes(htmlTemplate)))
        {
            // LoadOptions specifies that the source format is HTML.
            LoadOptions loadOptions = new LoadOptions { LoadFormat = LoadFormat.Html };
            templateDoc = new Document(htmlStream, loadOptions);
        }

        // Original data source is an array.
        string[] arrayData = new string[] { "Apple", "Banana", "Cherry" };

        // Convert the array to a collection type (List<string>) which the ReportingEngine can work with.
        List<string> collectionData = arrayData.ToList();

        // Create a ReportingEngine instance and build the report.
        ReportingEngine engine = new ReportingEngine();
        // The third argument is the name used in the template to reference the collection.
        engine.BuildReport(templateDoc, collectionData, "items");

        // Save the resulting document.
        templateDoc.Save("ReportFromHtml.docx");
    }
}
