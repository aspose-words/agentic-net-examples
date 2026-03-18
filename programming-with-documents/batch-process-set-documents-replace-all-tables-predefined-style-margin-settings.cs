using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing; // For LineStyle enum
using System.Drawing; // For Color

class Program
{
    static void Main()
    {
        // Folder that contains the source documents.
        string inputFolder = @"C:\Docs\Input";

        // Folder where the processed documents will be saved.
        string outputFolder = @"C:\Docs\Output";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Name of the predefined table style that will be applied to every table.
        const string predefinedStyleName = "MyBatchTableStyle";

        // Process each .docx file in the input folder.
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the document (uses the Document constructor – a lifecycle rule).
            Document doc = new Document(filePath);

            // -----------------------------------------------------------------
            // Create (or retrieve) the predefined table style in the current document.
            // -----------------------------------------------------------------
            TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, predefinedStyleName);

            // Define the visual appearance of the style.
            tableStyle.Borders.Color = Color.Blue;
            tableStyle.Borders.LineStyle = LineStyle.Single;
            tableStyle.Borders.LineWidth = 1.0; // points
            tableStyle.CellSpacing = 0; // no extra spacing between cells
            tableStyle.LeftPadding = 5;
            tableStyle.RightPadding = 5;
            tableStyle.TopPadding = 5;
            tableStyle.BottomPadding = 5;
            tableStyle.Shading.BackgroundPatternColor = Color.LightGray;

            // -----------------------------------------------------------------
            // Apply the style to every table in the document.
            // -----------------------------------------------------------------
            NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
            foreach (Table table in tables)
            {
                // Assign the predefined style.
                table.Style = tableStyle;

                // Apply margin‑like settings directly to the table.
                // These properties control the distance between the table and surrounding text.
                table.DistanceTop = 12;    // points
                table.DistanceBottom = 12; // points
                table.DistanceLeft = 12;   // points
                table.DistanceRight = 12;  // points

                // Optional: remove any left indent that might have been inherited.
                table.LeftIndent = 0;
            }

            // Convert style formatting to direct formatting so that the style is fully materialized.
            doc.ExpandTableStylesToDirectFormatting();

            // Save the modified document to the output folder (uses the Document.Save method – a lifecycle rule).
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
            doc.Save(outputPath);
        }
    }
}
