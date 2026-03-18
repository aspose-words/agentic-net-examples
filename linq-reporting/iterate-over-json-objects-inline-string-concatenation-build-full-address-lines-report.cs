using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Path to the JSON file that contains address data.
        string jsonPath = @"C:\Data\addresses.json";

        // Path to the output Word document.
        string outputPath = @"C:\Output\Report.docx";

        // -----------------------------------------------------------------
        // 1. Load the JSON data source.
        // -----------------------------------------------------------------
        // JsonDataSource parses the JSON file and makes its elements available
        // to the ReportingEngine. No custom options are required for this example.
        JsonDataSource jsonData = new JsonDataSource(jsonPath);

        // -----------------------------------------------------------------
        // 2. Create a template document programmatically.
        // -----------------------------------------------------------------
        // The template contains a single line that builds a full address using
        // inline string concatenation of the JSON fields.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Example JSON structure:
        // {
        //   "AddressLine1": "123 Main St",
        //   "City": "Springfield",
        //   "PostalCode": "12345"
        // }
        //
        // The template syntax <<[FieldName]>> inserts the value of a field.
        // The '+' operator concatenates strings.  Escape double quotes inside the
        // template string with backslashes.
        builder.Writeln(
            "Full Address: <<[AddressLine1] + \", \" + [City] + \" \" + [PostalCode]>>");

        // -----------------------------------------------------------------
        // 3. Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();

        // The third argument ("persons") is the name used to reference the data
        // source inside the template.  It can be any identifier; the template
        // does not need to prefix field names because we are using a single
        // top‑level JSON object.
        engine.BuildReport(template, jsonData, "persons");

        // -----------------------------------------------------------------
        // 4. Save the populated document.
        // -----------------------------------------------------------------
        template.Save(outputPath);
    }
}
