using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the PDF template that already contains a table.
        Document doc = new Document("Template.pdf");

        // Retrieve the first table in the document (adjust the index if needed).
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Create a custom table style that will hold our conditional formatting.
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomTableStyle");

        // Apply conditional formatting to the last column of the table:
        //   - Make the text bold.
        //   - Set a light gray background shading.
        customStyle.ConditionalStyles.LastColumn.Font.Bold = true;
        customStyle.ConditionalStyles.LastColumn.Shading.BackgroundPatternColor = Color.LightGray;

        // Assign the custom style to the table.
        table.Style = customStyle;

        // Enable the LastColumn conditional style so it takes effect.
        table.StyleOptions |= TableStyleOptions.LastColumn;

        // Save the modified document as a PDF.
        doc.Save("Result.pdf");
    }
}
