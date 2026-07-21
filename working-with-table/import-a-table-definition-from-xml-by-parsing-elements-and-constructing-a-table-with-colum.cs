using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

namespace TableFromXmlExample
{
    public class Program
    {
        public static void Main()
        {
            // Sample XML that defines a table, its style and column widths.
            const string xmlDefinition = @"
<TableDefinition>
    <StyleIdentifier>MediumShading1Accent1</StyleIdentifier>
    <Columns>
        <Column Width='100' />
        <Column Width='150' />
    </Columns>
    <Rows>
        <Row>
            <Cell>Item</Cell>
            <Cell>Quantity (kg)</Cell>
        </Row>
        <Row>
            <Cell>Apples</Cell>
            <Cell>20</Cell>
        </Row>
        <Row>
            <Cell>Bananas</Cell>
            <Cell>40</Cell>
        </Row>
        <Row>
            <Cell>Carrots</Cell>
            <Cell>50</Cell>
        </Row>
    </Rows>
</TableDefinition>";

            // Parse the XML.
            XDocument xDoc = XDocument.Parse(xmlDefinition);
            XElement root = xDoc.Root;

            // Extract the style identifier. If parsing fails, fall back to Normal style.
            string styleIdString = root.Element("StyleIdentifier")?.Value ?? "Normal";
            StyleIdentifier styleId = Enum.TryParse(styleIdString, out StyleIdentifier parsedStyle)
                ? parsedStyle
                : StyleIdentifier.Normal;

            // Extract column widths (in points).
            var columnWidths = root.Element("Columns")
                                   .Elements("Column")
                                   .Select(col => double.Parse(col.Attribute("Width")?.Value ?? "0"))
                                   .ToList();

            // Extract rows and cells.
            var rows = root.Element("Rows")
                           .Elements("Row")
                           .Select(r => r.Elements("Cell").Select(c => c.Value).ToList())
                           .ToList();

            // Create a new document and a builder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start the table.
            Table table = builder.StartTable();

            // Build the table based on the parsed XML.
            foreach (var rowCells in rows)
            {
                for (int colIndex = 0; colIndex < rowCells.Count; colIndex++)
                {
                    // Insert a new cell.
                    builder.InsertCell();

                    // Apply the column width for this cell if defined.
                    if (colIndex < columnWidths.Count)
                        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(columnWidths[colIndex]);

                    // Write the cell text.
                    builder.Write(rowCells[colIndex]);
                }

                // End the current row.
                builder.EndRow();
            }

            // Apply the style and fixed column widths after the table content is built.
            table.StyleIdentifier = styleId;
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            // Finish the table.
            builder.EndTable();

            // Define the output path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableFromXml.docx");

            // Save the document.
            doc.Save(outputPath);

            // Simple validation that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file at '{outputPath}'.");
        }
    }
}
