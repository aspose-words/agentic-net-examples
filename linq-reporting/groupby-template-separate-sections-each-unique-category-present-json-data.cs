using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsGroupByExample
{
    class Program
    {
        static void Main()
        {
            // Load the Word template that contains the Reporting Engine syntax.
            // The template should have a foreach block with a groupby clause, e.g.:
            // <<foreach [item in ds.Items groupby item.Category]>>
            //   <<[key]>>   // the current group key (category name)
            //   <<foreach [i in value]>>
            //       <<[i.Name]>>
            //   <</foreach>>
            // <</foreach>>
            Document template = new Document("Template.docx");

            // Load JSON data. The JSON file must contain an array named "Items"
            // where each element has at least a "Category" and a "Name" property.
            JsonDataSource jsonSource = new JsonDataSource("data.json");

            // Create the reporting engine and populate the template.
            ReportingEngine engine = new ReportingEngine();
            // The data source name ("ds") must match the name used in the template tags.
            engine.BuildReport(template, jsonSource, "ds");

            // Save the generated report.
            template.Save("Report.docx");
        }
    }
}
