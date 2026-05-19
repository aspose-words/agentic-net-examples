using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Sample XML that defines a table structure, column widths (in points) and optional cell styles.
        const string xml = @"
<TableDefinition>
    <Columns>
        <Column Width='100' />
        <Column Width='200' />
    </Columns>
    <Rows>
        <Row>
            <Cell Style='Bold'>Header 1</Cell>
            <Cell Style='Bold'>Header 2</Cell>
        </Row>
        <Row>
            <Cell>Row 1, Col 1</Cell>
            <Cell>Row 1, Col 2</Cell>
        </Row>
        <Row>
            <Cell>Row 2, Col 1</Cell>
            <Cell>Row 2, Col 2</Cell>
        </Row>
    </Rows>
</TableDefinition>";

        // Parse the XML.
        XDocument docXml = XDocument.Parse(xml);
        var columnWidths = docXml.Root
                                 .Element("Columns")
                                 .Elements("Column")
                                 .Select(c => double.Parse(c.Attribute("Width")!.Value))
                                 .ToList();

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building the table.
        Table table = builder.StartTable();

        // Iterate over each row defined in the XML.
        foreach (var rowElement in docXml.Root.Element("Rows")!.Elements("Row"))
        {
            // Process each cell in the current row.
            var cells = rowElement.Elements("Cell").ToList();
            for (int i = 0; i < cells.Count; i++)
            {
                // Insert a new cell.
                builder.InsertCell();

                // Apply column width based on the column index.
                if (i < columnWidths.Count)
                {
                    builder.CellFormat.Width = columnWidths[i];
                }

                // Apply simple style handling (currently only bold).
                bool isBold = (cells[i].Attribute("Style")?.Value ?? string.Empty).Equals("Bold", StringComparison.OrdinalIgnoreCase);
                builder.Font.Bold = isBold;

                // Write the cell text.
                builder.Write(cells[i].Value);

                // Reset bold formatting for subsequent cells.
                builder.Font.Bold = false;
            }

            // End the current row.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to a file in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OutputTable.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file at '{outputPath}'.");
        }

        // The program ends automatically; no user interaction required.
    }
}
