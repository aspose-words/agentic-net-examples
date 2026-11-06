using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class ImportTableFromXml
{
    public static void Main()
    {
        // Sample XML defining a table.
        string xml = @"
<TableDefinition>
    <Style>TableGrid</Style>
    <Columns>
        <Column Width='100'/>
        <Column Width='150'/>
        <Column Width='200'/>
    </Columns>
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

        // Parse XML.
        XDocument docXml = XDocument.Parse(xml);
        string styleName = docXml.Root.Element("Style")?.Value?.Trim() ?? "TableGrid";

        var columnWidths = docXml.Root.Element("Columns")
            .Elements("Column")
            .Select(c => (double)double.Parse(c.Attribute("Width")?.Value ?? "0"))
            .ToArray();

        var rows = docXml.Root.Element("Rows")
            .Elements("Row")
            .Select(r => r.Elements("Cell").Select(c => c.Value).ToArray())
            .ToArray();

        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build the table.
        builder.StartTable();

        for (int rowIndex = 0; rowIndex < rows.Length; rowIndex++)
        {
            string[] cells = rows[rowIndex];
            for (int colIndex = 0; colIndex < cells.Length; colIndex++)
            {
                builder.InsertCell();

                // Set cell text.
                builder.Writeln(cells[colIndex]);

                // Apply column width if defined.
                if (colIndex < columnWidths.Length && columnWidths[colIndex] > 0)
                {
                    Cell cell = (Cell)builder.CurrentParagraph.ParentNode;
                    cell.CellFormat.PreferredWidth = PreferredWidth.FromPoints(columnWidths[colIndex]);
                }
            }
            builder.EndRow();
        }

        builder.EndTable();

        // Apply table style.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];
        // Try to map the style name to a built‑in identifier.
        if (Enum.TryParse(styleName, true, out StyleIdentifier styleId) && Enum.IsDefined(typeof(StyleIdentifier), styleId))
        {
            table.StyleIdentifier = styleId;
        }
        else
        {
            table.StyleName = styleName;
        }

        // Save the document.
        string outputPath = "OutputTable.docx";
        doc.Save(outputPath);

        // Verify the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create output file: {outputPath}");
    }
}
