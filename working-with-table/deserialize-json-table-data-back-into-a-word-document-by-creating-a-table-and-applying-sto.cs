using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

namespace AsposeWordsTableFromJson
{
    // Represents the whole table structure.
    public class TableData
    {
        public List<RowData> Rows { get; set; }
    }

    // Represents a single row.
    public class RowData
    {
        public List<CellData> Cells { get; set; }
    }

    // Represents a single cell with optional formatting.
    public class CellData
    {
        public string Text { get; set; }

        // Optional width in points.
        public double? Width { get; set; }

        // Optional padding values in points.
        public double? LeftPadding { get; set; }
        public double? RightPadding { get; set; }
        public double? TopPadding { get; set; }
        public double? BottomPadding { get; set; }

        // Optional background color in hex format, e.g. "#FFCC00".
        public string BackgroundColor { get; set; }

        // Optional vertical alignment: "Top", "Center", "Bottom".
        public string VerticalAlignment { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Sample JSON describing a table with two rows and two columns.
            string json = @"
{
  ""Rows"": [
    {
      ""Cells"": [
        {
          ""Text"": ""Header 1"",
          ""BackgroundColor"": ""#D9E1F2"",
          ""VerticalAlignment"": ""Center"",
          ""Width"": 150,
          ""LeftPadding"": 5,
          ""RightPadding"": 5,
          ""TopPadding"": 2,
          ""BottomPadding"": 2
        },
        {
          ""Text"": ""Header 2"",
          ""BackgroundColor"": ""#D9E1F2"",
          ""VerticalAlignment"": ""Center"",
          ""Width"": 150,
          ""LeftPadding"": 5,
          ""RightPadding"": 5,
          ""TopPadding"": 2,
          ""BottomPadding"": 2
        }
      ]
    },
    {
      ""Cells"": [
        {
          ""Text"": ""Row 1, Cell 1"",
          ""BackgroundColor"": ""#FFFFFF"",
          ""VerticalAlignment"": ""Top""
        },
        {
          ""Text"": ""Row 1, Cell 2"",
          ""BackgroundColor"": ""#FFFFFF"",
          ""VerticalAlignment"": ""Top""
        }
      ]
    }
  ]
}";
            // Deserialize JSON into TableData object.
            TableData tableData = JsonConvert.DeserializeObject<TableData>(json);
            if (tableData == null || tableData.Rows == null)
                throw new InvalidOperationException("Failed to deserialize table data.");

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start building the table.
            Table table = builder.StartTable();

            foreach (RowData rowData in tableData.Rows)
            {
                if (rowData?.Cells == null)
                    continue;

                foreach (CellData cellData in rowData.Cells)
                {
                    // Insert a new cell.
                    builder.InsertCell();

                    // Apply optional formatting.
                    if (cellData.Width.HasValue)
                        builder.CellFormat.Width = cellData.Width.Value;

                    if (cellData.LeftPadding.HasValue)
                        builder.CellFormat.LeftPadding = cellData.LeftPadding.Value;
                    if (cellData.RightPadding.HasValue)
                        builder.CellFormat.RightPadding = cellData.RightPadding.Value;
                    if (cellData.TopPadding.HasValue)
                        builder.CellFormat.TopPadding = cellData.TopPadding.Value;
                    if (cellData.BottomPadding.HasValue)
                        builder.CellFormat.BottomPadding = cellData.BottomPadding.Value;

                    if (!string.IsNullOrEmpty(cellData.BackgroundColor))
                    {
                        Color bg = ColorTranslator.FromHtml(cellData.BackgroundColor);
                        builder.CellFormat.Shading.BackgroundPatternColor = bg;
                    }

                    if (!string.IsNullOrEmpty(cellData.VerticalAlignment))
                    {
                        if (Enum.TryParse(cellData.VerticalAlignment, out CellVerticalAlignment vAlign))
                            builder.CellFormat.VerticalAlignment = vAlign;
                    }

                    // Write the cell text.
                    builder.Write(cellData.Text ?? string.Empty);
                }

                // End the current row.
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OutputTable.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("The output document was not saved.", outputPath);
        }
    }
}
