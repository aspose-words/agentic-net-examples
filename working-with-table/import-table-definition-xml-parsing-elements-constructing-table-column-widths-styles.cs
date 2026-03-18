using System;
using System.IO;
using System.Xml;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class XmlTableImporter
{
    static void Main()
    {
        // Path to the XML definition file.
        const string xmlPath = "TableDefinition.xml";

        // Load the XML document.
        XmlDocument xmlDoc = new XmlDocument();
        xmlDoc.Load(xmlPath);

        // Create a new empty Word document.
        Document doc = new Document();

        // Start a new table using DocumentBuilder.
        DocumentBuilder builder = new DocumentBuilder(doc);
        Table table = builder.StartTable();

        // -----------------------------------------------------------------
        // Parse column widths (comma‑separated list of points).
        // Example: <ColumnWidths>100,150,200</ColumnWidths>
        // -----------------------------------------------------------------
        XmlNode widthsNode = xmlDoc.SelectSingleNode("//ColumnWidths");
        double[] columnWidths = Array.Empty<double>();
        if (widthsNode != null && !string.IsNullOrWhiteSpace(widthsNode.InnerText))
        {
            string[] parts = widthsNode.InnerText.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            columnWidths = new double[parts.Length];
            for (int i = 0; i < parts.Length; i++)
            {
                if (double.TryParse(parts[i], out double w))
                    columnWidths[i] = w;
                else
                    columnWidths[i] = 0; // fallback to default width
            }
        }

        // -----------------------------------------------------------------
        // Parse optional table style identifier.
        // Example: <Style>MediumShading1Accent1</Style>
        // -----------------------------------------------------------------
        XmlNode styleNode = xmlDoc.SelectSingleNode("//Style");
        if (styleNode != null && !string.IsNullOrWhiteSpace(styleNode.InnerText))
        {
            if (Enum.TryParse(styleNode.InnerText, out StyleIdentifier styleId))
                table.StyleIdentifier = styleId;
        }

        // -----------------------------------------------------------------
        // Parse rows and cells.
        // Expected XML:
        // <Rows>
        //   <Row>
        //     <Cell>First cell text</Cell>
        //     <Cell>Second cell text</Cell>
        //   </Row>
        //   ...
        // </Rows>
        // -----------------------------------------------------------------
        XmlNodeList rowNodes = xmlDoc.SelectNodes("//Rows/Row");
        foreach (XmlNode rowNode in rowNodes)
        {
            // Ensure we have at least one cell before writing content.
            builder.InsertCell();

            XmlNodeList cellNodes = rowNode.SelectNodes("Cell");
            for (int cellIndex = 0; cellIndex < cellNodes.Count; cellIndex++)
            {
                // For the first cell of the row we already inserted one above.
                if (cellIndex > 0)
                    builder.InsertCell();

                // Write the cell text.
                builder.Write(cellNodes[cellIndex].InnerText ?? string.Empty);

                // Apply column width if we have a definition for this column.
                if (cellIndex < columnWidths.Length && columnWidths[cellIndex] > 0)
                {
                    // Width is expressed in points.
                    builder.CellFormat.Width = columnWidths[cellIndex];
                }
            }

            // End the current row.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // -----------------------------------------------------------------
        // Save the document.
        // -----------------------------------------------------------------
        const string outputPath = "ImportedTable.docx";
        doc.Save(outputPath, SaveFormat.Docx);
    }
}
