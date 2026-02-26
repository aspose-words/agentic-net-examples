using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables; // Added for Table class

namespace AsposeWordsLinqReportingDemo
{
    // Simple data model used as the data source for the report.
    public class Item
    {
        public string? Name { get; set; } // Made nullable to silence CS8618 warning
        public bool IsActive { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a WORDML template that contains a table with a conditional row.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Create a table with a header row.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Item Name");
            builder.EndRow();

            // Insert a data row that will be repeated for each item and displayed only when the condition is true.
            // The LINQ Reporting Engine uses the syntax <<foreach [item]>> ... <<endforeach>> to repeat rows.
            // Inside the row we use <<if [item.IsActive]>> to conditionally output the cell content.
            builder.InsertCell();
            builder.Write("<<foreach [item]>>");               // start repeating this row for each item
            builder.Write("<<if [item.IsActive]>>");          // condition start
            builder.Write("<<[item.Name]>>");                 // field output
            builder.Write("<<endif>>");                       // condition end
            builder.Write("<<endforeach>>");                 // end of repetition for this row
            builder.EndRow();

            builder.EndTable();

            // Save the template to disk (WORDML format is just a .docx file with the appropriate tags).
            const string templatePath = "ConditionalTableTemplate.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report using a list of items.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Prepare sample data.
            List<Item> items = new List<Item>
            {
                new Item { Name = "Alpha",   IsActive = true  },
                new Item { Name = "Beta",    IsActive = false },
                new Item { Name = "Gamma",   IsActive = true  },
                new Item { Name = "Delta",   IsActive = false },
                new Item { Name = "Epsilon", IsActive = true  }
            };

            // The ReportingEngine expects the data source name to be referenced in the template.
            // In the template we used the prefix "item" (e.g., <<[item.Name]>>), so we pass that name.
            ReportingEngine engine = new ReportingEngine();

            // Build the report. The engine will repeat the table row for each item in the list
            // and will include the row only when item.IsActive evaluates to true.
            engine.BuildReport(reportDoc, items, "item");

            // -----------------------------------------------------------------
            // 3. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "ConditionalTableReport.docx";
            reportDoc.Save(outputPath);
        }
    }
}
