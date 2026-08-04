using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

namespace TableFromXmlExample
{
    public class Program
    {
        public static void Main()
        {
            // XML definition of the table.
            const string xml = @"
<TableDefinition>
    <Columns>
        <Column Width='120' />
        <Column Width='200' />
        <Column Width='150' />
    </Columns>
    <Rows>
        <Row>
            <Cell>Product</Cell>
            <Cell>Quantity</Cell>
            <Cell>Price</Cell>
        </Row>
        <Row>
            <Cell>Apples</Cell>
            <Cell>20</Cell>
            <Cell>$1.50</Cell>
        </Row>
        <Row>
            <Cell>Bananas</Cell>
            <Cell>35</Cell>
            <Cell>$0.80</Cell>
        </Row>
    </Rows>
    <Style>
        <HeaderShadingColor>#D3D3D3</HeaderShadingColor>
    </Style>
</TableDefinition>";

            // Parse the XML.
            XDocument xDoc = XDocument.Parse(xml);
            var columnWidths = xDoc.Root
                                   .Element("Columns")
                                   .Elements("Column")
                                   .Select(c => (double)c.Attribute("Width"))
                                   .ToArray();

            var rows = xDoc.Root
                           .Element("Rows")
                           .Elements("Row")
                           .Select(r => r.Elements("Cell").Select(c => c.Value).ToArray())
                           .ToArray();

            string headerShadingHex = (string)xDoc.Root.Element("Style")?.Element("HeaderShadingColor");
            Color headerShading = Color.Empty;
            if (!string.IsNullOrEmpty(headerShadingHex))
            {
                headerShading = ColorTranslator.FromHtml(headerShadingHex);
            }

            // Create a new document and a builder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start the table.
            Table table = builder.StartTable();

            // Build rows.
            for (int rowIndex = 0; rowIndex < rows.Length; rowIndex++)
            {
                string[] cells = rows[rowIndex];
                for (int colIndex = 0; colIndex < cells.Length; colIndex++)
                {
                    // Insert a new cell.
                    builder.InsertCell();

                    // Set the column width.
                    builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(columnWidths[colIndex]);

                    // Apply header shading to the first row.
                    if (rowIndex == 0 && headerShading != Color.Empty)
                    {
                        builder.CellFormat.Shading.BackgroundPatternColor = headerShading;
                    }

                    // Write the cell text.
                    builder.Write(cells[colIndex]);
                }

                // End the current row.
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Save the document.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);
            string outputPath = Path.Combine(outputDir, "TableFromXml.docx");
            doc.Save(outputPath);

            // Simple validation that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not created.");

            // The program ends automatically; no user interaction required.
        }
    }
}
