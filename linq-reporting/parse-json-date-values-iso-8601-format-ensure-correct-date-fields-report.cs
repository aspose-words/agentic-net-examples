using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the Word template that contains reporting engine placeholders.
        Document doc = new Document("Template.docx");

        // Set up JSON loading options to explicitly recognize ISO‑8601 date formats.
        // This ensures that date strings like "2023-04-15T13:45:30Z" are parsed as DateTime values.
        JsonDataLoadOptions jsonOptions = new JsonDataLoadOptions
        {
            ExactDateTimeParseFormats = new List<string>
            {
                "yyyy-MM-ddTHH:mm:ssK", // Full ISO‑8601 with timezone (e.g., Z or +02:00)
                "yyyy-MM-ddTHH:mm:ss",  // Without timezone
                "yyyy-MM-dd"            // Date only
            }
        };

        // Create a JSON data source using the options defined above.
        JsonDataSource dataSource = new JsonDataSource("Data.json", jsonOptions);

        // ReportingEngine is an instance class; create an instance before calling BuildReport.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "records");

        // Save the generated document.
        doc.Save("Report.docx");
    }
}
