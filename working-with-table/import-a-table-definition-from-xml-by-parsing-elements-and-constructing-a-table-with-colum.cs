using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class ImportTableFromXml
{
    public static void Main()
    {
        // Sample XML defining a table with column widths and simple style.
        string xmlContent = @"
<Table>
  <Columns>
    <Column Width='100' />
    <Column Width='150' />
  </Columns>
  <Rows>
    <Row>
      <Cell>Header 1</Cell>
      <Cell>Header 2</Cell>
    </Row>
    <Row>
      <Cell>Data 1</Cell>
      <Cell>Data 2</Cell>
    </Row>
  </Rows>
  <Style>
    <BorderColor>#0000FF</BorderColor>
    <ShadingColor>#FFFFCC</ShadingColor>
  </Style>
</Table>";

        // Parse the XML.
        XDocument xDoc = XDocument.Parse(xmlContent);
        XElement tableElem = xDoc.Element("Table");
        if (tableElem == null) throw new Exception("Invalid XML: missing Table element.");

        // Extract column widths.
        List<double> columnWidths = tableElem.Element("Columns")?
            .Elements("Column")
            .Select(c => (double)double.Parse(c.Attribute("Width")?.Value ?? "0"))
            .ToList() ?? new List<double>();

        // Extract rows and cells.
        List<List<string>> rows = tableElem.Element("Rows")?
            .Elements("Row")
            .Select(r => r.Elements("Cell").Select(c => c.Value).ToList())
            .ToList() ?? new List<List<string>>();

        // Extract style information.
        XElement styleElem = tableElem.Element("Style");
        Color borderColor = Color.Black;
        Color shadingColor = Color.Empty;
        if (styleElem != null)
        {
            string borderHex = styleElem.Element("BorderColor")?.Value;
            if (!string.IsNullOrEmpty(borderHex))
                borderColor = ColorTranslator.FromHtml(borderHex);

            string shadingHex = styleElem.Element("ShadingColor")?.Value;
            if (!string.IsNullOrEmpty(shadingHex))
                shadingColor = ColorTranslator.FromHtml(shadingHex);
        }

        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building the table.
        builder.StartTable();

        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            List<string> cells = rows[rowIndex];
            for (int colIndex = 0; colIndex < cells.Count; colIndex++)
            {
                // Insert a new cell.
                builder.InsertCell();

                // Write cell text.
                builder.Write(cells[colIndex]);

                // Apply column width if defined.
                if (colIndex < columnWidths.Count && columnWidths[colIndex] > 0)
                {
                    builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(columnWidths[colIndex]);
                }

                // Apply shading if defined.
                if (shadingColor != Color.Empty)
                {
                    builder.CellFormat.Shading.BackgroundPatternColor = shadingColor;
                }

                // Apply border color to all sides.
                builder.CellFormat.Borders.Color = borderColor;
            }

            // End the current row.
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "OutputTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");

        // Optional: clean up resources (handled by .NET runtime).
    }
}
