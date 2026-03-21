using System;
using System.Xml;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class XmlTableImporter
{
    static void Main()
    {
        // Embedded XML definition.
        const string xmlContent = @"
<TableDefinition>
    <ColumnWidths>100,150,200</ColumnWidths>
    <Style>MediumShading1Accent1</Style>
    <Rows>
        <Row>
            <Cell>Header 1</Cell>
            <Cell>Header 2</Cell>
            <Cell>Header 3</Cell>
        </Row>
        <Row>
            <Cell>Data 1</Cell>
            <Cell>Data 2</Cell>
            <Cell>Data 3</Cell>
        </Row>
    </Rows>
</TableDefinition>";

        // Load the XML document from the embedded string.
        XmlDocument xmlDoc = new XmlDocument();
        xmlDoc.LoadXml(xmlContent);

        // Create a new empty Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Table table = builder.StartTable();

        // Parse column widths (comma‑separated list of points).
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
                    columnWidths[i] = 0;
            }
        }

        // Store optional table style identifier for later application.
        StyleIdentifier? styleId = null;
        XmlNode styleNode = xmlDoc.SelectSingleNode("//Style");
        if (styleNode != null && !string.IsNullOrWhiteSpace(styleNode.InnerText))
        {
            if (Enum.TryParse(styleNode.InnerText, out StyleIdentifier parsedStyle))
                styleId = parsedStyle;
        }

        // Parse rows and cells.
        XmlNodeList rowNodes = xmlDoc.SelectNodes("//Rows/Row");
        foreach (XmlNode rowNode in rowNodes)
        {
            XmlNodeList cellNodes = rowNode.SelectNodes("Cell");
            for (int cellIndex = 0; cellIndex < cellNodes.Count; cellIndex++)
            {
                // Apply column width before inserting the cell, if defined.
                if (cellIndex < columnWidths.Length && columnWidths[cellIndex] > 0)
                    builder.CellFormat.Width = columnWidths[cellIndex];
                else
                    builder.CellFormat.Width = 0; // reset to default when not specified

                builder.InsertCell();
                builder.Write(cellNodes[cellIndex].InnerText ?? string.Empty);
            }
            builder.EndRow();
        }

        // Apply the style after the table has at least one row.
        if (styleId.HasValue)
            table.StyleIdentifier = styleId.Value;

        // Finish the table.
        builder.EndTable();

        // Save the document.
        const string outputPath = "ImportedTable.docx";
        doc.Save(outputPath, SaveFormat.Docx);
    }
}
