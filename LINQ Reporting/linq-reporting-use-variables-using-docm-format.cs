using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReportingExample
{
    // Simple data source class for LINQ Reporting.
    public class ReportData
    {
        public string Title { get; set; }
        public string Author { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new blank document.
            Document doc = new Document();

            // 2. Build the template content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a DOCVARIABLE field that will display the value of a document variable.
            FieldDocVariable companyField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
            companyField.VariableName = "CompanyName";

            // Add a line break after the field.
            builder.Writeln();

            // Insert LINQ Reporting placeholders that will be replaced from the data source.
            // The syntax <<[data.Title]>> and <<[data.Author]>> will be processed by ReportingEngine.
            builder.Writeln("Report Title: <<[data.Title]>>");
            builder.Writeln("Report Author: <<[data.Author]>>");

            // 3. Add a document variable that the DOCVARIABLE field will read.
            doc.Variables.Add("CompanyName", "Acme Corporation");

            // 4. Prepare the data source instance.
            ReportData data = new ReportData
            {
                Title = "Annual Financial Summary",
                Author = "John Doe"
            };

            // 5. Build the report using LINQ ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The second parameter is the data source object, the third is the name used in the template.
            engine.BuildReport(doc, data, "data");

            // 6. Save the resulting document in DOCM format (macro‑enabled Word document).
            doc.Save("LinqReport.docm", SaveFormat.Docm);
        }
    }
}
