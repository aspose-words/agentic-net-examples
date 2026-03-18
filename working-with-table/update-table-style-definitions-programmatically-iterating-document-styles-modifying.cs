using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableStyleUpdate
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing document (replace with your actual file path)
            Document doc = new Document("Input.docx");

            // Iterate through all styles in the document
            foreach (Style style in doc.Styles)
            {
                // Process only table styles
                if (style.Type == StyleType.Table)
                {
                    // Cast the generic Style to TableStyle to access table‑specific properties
                    TableStyle tableStyle = (TableStyle)style;

                    // Example modifications:
                    // Set the number of rows to be banded (odd/even row shading)
                    tableStyle.RowStripe = 2;

                    // Set spacing between cells
                    tableStyle.CellSpacing = 4.0;

                    // Change the background color of the table cells
                    tableStyle.Shading.BackgroundPatternColor = Color.LightYellow;

                    // Set a uniform border color and style for the table
                    tableStyle.Borders.Color = Color.DarkBlue;
                    tableStyle.Borders.LineStyle = LineStyle.Single;
                    tableStyle.Borders.LineWidth = 1.0;

                    // Adjust vertical alignment of cells
                    tableStyle.VerticalAlignment = CellVerticalAlignment.Center;

                    // Optionally modify conditional styles (e.g., first row)
                    tableStyle.ConditionalStyles.FirstRow.Shading.BackgroundPatternColor = Color.LightBlue;
                    tableStyle.ConditionalStyles.FirstRow.Font.Bold = true;
                }
            }

            // If you need the style changes to be reflected as direct formatting on existing tables,
            // expand the table styles to direct formatting.
            doc.ExpandTableStylesToDirectFormatting();

            // Save the modified document (replace with your desired output path)
            doc.Save("Output.docx");
        }
    }
}
