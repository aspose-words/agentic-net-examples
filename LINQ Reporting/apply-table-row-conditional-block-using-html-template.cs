using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // HTML template that contains a simple table.
        string html = @"
<table>
    <tr><td>Header 1</td><td>Header 2</td></tr>
    <tr><td>Row 1, Col 1</td><td>Row 1, Col 2</td></tr>
    <tr><td>Row 2, Col 1</td><td>Row 2, Col 2</td></tr>
</table>";

        // Insert the HTML into the document. The table becomes a native Aspose.Words Table node.
        builder.InsertHtml(html);

        // Retrieve the first (and only) table that was just inserted.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Create a custom table style that will hold conditional formatting.
        TableStyle conditionalStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyConditionalStyle");

        // Apply a background color to the last row of the table via a conditional style.
        conditionalStyle.ConditionalStyles[ConditionalStyleType.LastRow].Shading.BackgroundPatternColor = Color.LightGray;

        // Assign the custom style to the table.
        table.Style = conditionalStyle;

        // Enable the LastRow conditional formatting for this table.
        table.StyleOptions |= TableStyleOptions.LastRow;

        // Save the resulting document.
        doc.Save("ConditionalRowTable.docx");
    }
}
