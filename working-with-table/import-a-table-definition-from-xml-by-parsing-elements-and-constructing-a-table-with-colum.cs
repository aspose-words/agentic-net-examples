using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableImportFromXml
{
    public class Program
    {
        public static void Main()
        {
            // Sample XML defining a table structure.
            string xmlContent = @"
<TableDefinition>
  <Columns>
    <Column Index='0' Width='100' />
    <Column Index='1' Width='150' />
    <Column Index='2' Width='120' />
  </Columns>
  <Rows>
    <Row>
      <Cell ColumnIndex='0' Text='Header 1' Style='Header' />
      <Cell ColumnIndex='1' Text='Header 2' Style='Header' />
      <Cell ColumnIndex='2' Text='Header 3' Style='Header' />
    </Row>
    <Row>
      <Cell ColumnIndex='0' Text='Row1 Col1' />
      <Cell ColumnIndex='1' Text='Row1 Col2' />
      <Cell ColumnIndex='2' Text='Row1 Col3' />
    </Row>
    <Row>
      <Cell ColumnIndex='0' Text='Row2 Col1' />
      <Cell ColumnIndex='1' Text='Row2 Col2' />
      <Cell ColumnIndex='2' Text='Row2 Col3' />
    </Row>
  </Rows>
</TableDefinition>";

            // Parse XML.
            XDocument xDoc = XDocument.Parse(xmlContent);
            var columnElements = xDoc.Root.Element("Columns").Elements("Column");
            var rowElements = xDoc.Root.Element("Rows").Elements("Row");

            // Store column widths keyed by index.
            var columnWidths = new Dictionary<int, double>();
            foreach (var col in columnElements)
            {
                int index = (int)col.Attribute("Index");
                double width = double.Parse(col.Attribute("Width").Value);
                columnWidths[index] = width;
            }

            // Create a new document and builder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start building the table.
            builder.StartTable();

            // Iterate through each row defined in XML.
            foreach (var row in rowElements)
            {
                // Get cells ordered by ColumnIndex to ensure correct column order.
                var cells = row.Elements("Cell")
                               .OrderBy(c => (int)c.Attribute("ColumnIndex"))
                               .ToList();

                foreach (var cell in cells)
                {
                    int colIndex = (int)cell.Attribute("ColumnIndex");
                    string text = (string)cell.Attribute("Text") ?? string.Empty;
                    string style = (string)cell.Attribute("Style") ?? string.Empty;

                    // Apply column width.
                    if (columnWidths.TryGetValue(colIndex, out double width))
                    {
                        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(width);
                    }

                    // Apply simple header style (shading) if requested.
                    if (style.Equals("Header", StringComparison.OrdinalIgnoreCase))
                    {
                        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
                    }
                    else
                    {
                        // Reset shading for non‑header cells.
                        builder.CellFormat.Shading.BackgroundPatternColor = Color.Empty;
                    }

                    // Insert the cell and write its text.
                    builder.InsertCell();
                    builder.Write(text);
                }

                // End the current row.
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Save the document.
            string outputPath = "OutputTable.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
            {
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
            }
        }
    }
}
