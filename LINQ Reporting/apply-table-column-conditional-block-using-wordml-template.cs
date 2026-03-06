using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the WORDML template.
        Document doc = new Document("Template.docx");

        // Locate the first table in the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("No table found in the document.");

        // Create a custom table style (or reuse an existing one).
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomStyle");

        // Apply a conditional style to the last column: make the text bold and set a background color.
        ConditionalStyle lastColumnStyle = customStyle.ConditionalStyles.LastColumn;
        lastColumnStyle.Font.Bold = true;
        lastColumnStyle.Shading.BackgroundPatternColor = System.Drawing.Color.LightYellow;

        // Assign the custom style to the table.
        table.Style = customStyle;

        // Enable the LastColumn conditional formatting for this table.
        table.StyleOptions |= TableStyleOptions.LastColumn;

        // Save the modified document.
        doc.Save("Result.docx");
    }
}
