using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Settings;

namespace AsposeWordsConditionalRowExample
{
    class Program
    {
        static void Main()
        {
            // Load the RTF template that already contains a table.
            Document doc = new Document("Template.rtf");

            // Assume the first table in the document is the target.
            Table table = doc.FirstSection.Body.Tables[0];

            // Create a custom table style that will hold conditional formatting.
            TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyConditionalStyle");

            // Define a conditional style for the last row (e.g., yellow background).
            ConditionalStyle lastRowStyle = customStyle.ConditionalStyles[ConditionalStyleType.LastRow];
            lastRowStyle.Shading.BackgroundPatternColor = Color.Yellow;
            lastRowStyle.Font.Bold = true; // optional additional formatting

            // Apply the custom style to the table.
            table.Style = customStyle;

            // Enable the conditional formatting for the last row.
            table.StyleOptions |= TableStyleOptions.LastRow;

            // Save the modified document.
            doc.Save("Result.docx");
        }
    }
}
