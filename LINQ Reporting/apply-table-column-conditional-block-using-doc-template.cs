using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Reporting;
using System.Drawing;

class ApplyConditionalColumnStyle
{
    static void Main()
    {
        // Load the DOC template that already contains a table.
        Document doc = new Document("Template.doc");

        // Assume the template has at least one table; get the first one.
        Table table = doc.FirstSection.Body.Tables[0];

        // Ensure the table has a style assigned. If not, assign a built‑in style.
        if (table.Style == null)
        {
            table.StyleIdentifier = StyleIdentifier.TableGrid;
        }

        // Enable the conditional formatting for the last column.
        // This tells Word to apply the style defined in the LastColumn conditional style.
        table.StyleOptions |= TableStyleOptions.LastColumn;

        // Access the conditional style for the last column via the table's style.
        // If the table uses a custom style, we need to retrieve it; otherwise we work with the built‑in style.
        TableStyle tableStyle = (TableStyle)doc.Styles[table.StyleName];
        ConditionalStyle lastColumnStyle = tableStyle.ConditionalStyles.LastColumn;

        // Apply desired formatting to the last column conditional style.
        // Example: set a light gray background and make the text bold.
        lastColumnStyle.Shading.BackgroundPatternColor = Color.LightGray;
        lastColumnStyle.Font.Bold = true;

        // (Optional) Populate the table with data using the ReportingEngine.
        // This step demonstrates how a template can be merged with a data source.
        var data = new
        {
            Items = new[]
            {
                new { Name = "Apple",  Quantity = 10 },
                new { Name = "Banana", Quantity = 20 },
                new { Name = "Carrot", Quantity = 30 }
            }
        };
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, data, "ds");

        // Save the modified document.
        doc.Save("Result.docx");
    }
}
