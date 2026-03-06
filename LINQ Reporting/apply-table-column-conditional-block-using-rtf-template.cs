using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyConditionalColumnStyle
{
    static void Main()
    {
        // Load the RTF template.
        Document doc = new Document("Template.rtf");

        // Find the first table in the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("No table found in the document.");

        // Create a custom table style.
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomStyle");

        // Apply conditional formatting to the first column of the table style.
        // Example: set background shading and make the text bold.
        ConditionalStyle firstColumnStyle = customStyle.ConditionalStyles[ConditionalStyleType.FirstColumn];
        firstColumnStyle.Shading.BackgroundPatternColor = Color.LightYellow;
        firstColumnStyle.Font.Bold = true;

        // Assign the custom style to the table.
        table.Style = customStyle;

        // Enable the conditional style for the first column.
        table.StyleOptions = TableStyleOptions.FirstColumn;

        // Save the modified document.
        doc.Save("Result.docx");
    }
}
