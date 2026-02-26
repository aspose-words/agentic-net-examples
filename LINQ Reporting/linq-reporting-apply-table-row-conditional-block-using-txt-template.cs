using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model for the report.
    public class Item
    {
        public string Name { get; set; }
        public double Value { get; set; }
        public bool IsVisible { get; set; }
    }

    // Root object that will be passed to the ReportingEngine.
    public class ReportData
    {
        public List<Item> Items { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Prepare the data source.
            // -------------------------------------------------
            var data = new ReportData
            {
                Items = new List<Item>
                {
                    new Item { Name = "Alpha",   Value = 123.45, IsVisible = true  },
                    new Item { Name = "Beta",    Value = 678.90, IsVisible = false }, // This row will be omitted.
                    new Item { Name = "Gamma",   Value = 111.22, IsVisible = true  }
                }
            };

            // -------------------------------------------------
            // 2. Load the TXT template.
            //    The template contains a table row with a conditional block:
            //    <<foreach [Items]>>
            //        <<if [IsVisible]>>
            //            <<[Name]>>    <<[Value]>>
            //        <<endif>>
            //    <<endfor>>
            // -------------------------------------------------
            Document doc = new Document("Template.txt");

            // -------------------------------------------------
            // 3. Build the report using ReportingEngine.
            // -------------------------------------------------
            var engine = new ReportingEngine
            {
                // Remove empty paragraphs that may appear after conditional blocks are removed.
                Options = ReportBuildOptions.RemoveEmptyParagraphs | ReportBuildOptions.AllowMissingMembers
            };

            // The second overload allows us to reference the data source itself in the template via the name "ds".
            engine.BuildReport(doc, data, "ds");

            // -------------------------------------------------
            // 4. Save the populated document.
            // -------------------------------------------------
            doc.Save("Report.docx");
        }
    }
}
