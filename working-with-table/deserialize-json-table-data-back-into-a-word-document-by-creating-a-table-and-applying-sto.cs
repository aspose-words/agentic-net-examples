using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    // Classes that represent the JSON structure.
    public class TableData
    {
        public RowData[] Rows { get; set; }
        public TableFormatting TableFormatting { get; set; }
    }

    public class RowData
    {
        public CellData[] Cells { get; set; }
    }

    public class CellData
    {
        public string Text { get; set; }
        public string ShadingColor { get; set; } // Hex color, e.g. "#D9E1F2"
    }

    public class TableFormatting
    {
        public string BorderColor { get; set; }   // Hex color
        public double BorderLineWidth { get; set; } // In points
    }

    public static void Main()
    {
        // Sample JSON that stores table content and simple formatting.
        string json = @"
        {
            ""Rows"": [
                {
                    ""Cells"": [
                        { ""Text"": ""Header 1"", ""ShadingColor"": ""#D9E1F2"" },
                        { ""Text"": ""Header 2"", ""ShadingColor"": ""#D9E1F2"" }
                    ]
                },
                {
                    ""Cells"": [
                        { ""Text"": ""Row1Col1"", ""ShadingColor"": ""#FFFFFF"" },
                        { ""Text"": ""Row1Col2"", ""ShadingColor"": ""#FFFFFF"" }
                    ]
                },
                {
                    ""Cells"": [
                        { ""Text"": ""Row2Col1"", ""ShadingColor"": ""#FFFFFF"" },
                        { ""Text"": ""Row2Col2"", ""ShadingColor"": ""#FFFFFF"" }
                    ]
                }
            ],
            ""TableFormatting"": {
                ""BorderColor"": ""#0000FF"",
                ""BorderLineWidth"": 1.0
            }
        }";

        // Deserialize JSON into C# objects.
        TableData tableData = JsonConvert.DeserializeObject<TableData>(json);
        if (tableData == null || tableData.Rows == null)
            throw new InvalidOperationException("Failed to deserialize table data.");

        // Create a new blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build the table using DocumentBuilder.StartTable workflow.
        builder.StartTable();

        foreach (RowData row in tableData.Rows)
        {
            foreach (CellData cell in row.Cells)
            {
                // Insert a new cell.
                builder.InsertCell();

                // Write the cell text.
                builder.Write(cell.Text);

                // Apply shading to the current cell.
                Cell currentCell = builder.CurrentParagraph.ParentNode as Cell;
                if (currentCell != null && !string.IsNullOrEmpty(cell.ShadingColor))
                {
                    currentCell.CellFormat.Shading.BackgroundPatternColor =
                        ColorTranslator.FromHtml(cell.ShadingColor);
                }
            }

            // End the current row.
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Apply table border formatting if specified.
        Table table = doc.GetChildNodes(NodeType.Table, true)[0] as Table;
        if (table != null && tableData.TableFormatting != null)
        {
            Color borderColor = ColorTranslator.FromHtml(tableData.TableFormatting.BorderColor);
            double lineWidth = tableData.TableFormatting.BorderLineWidth;

            // Apply the same border to all sides.
            table.SetBorder(BorderType.Left, LineStyle.Single, lineWidth, borderColor, true);
            table.SetBorder(BorderType.Right, LineStyle.Single, lineWidth, borderColor, true);
            table.SetBorder(BorderType.Top, LineStyle.Single, lineWidth, borderColor, true);
            table.SetBorder(BorderType.Bottom, LineStyle.Single, lineWidth, borderColor, true);
            table.SetBorder(BorderType.Horizontal, LineStyle.Single, lineWidth, borderColor, true);
            table.SetBorder(BorderType.Vertical, LineStyle.Single, lineWidth, borderColor, true);
        }

        // Save the document.
        string outputPath = "OutputTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
