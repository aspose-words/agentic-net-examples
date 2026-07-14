using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeTableImportExample
{
    public class Program
    {
        public static void Main()
        {
            // Sample XML that defines a table: column widths, rows, and an optional style.
            const string tableXml = @"
<TableDefinition>
    <Columns>
        <Column Width='100' />
        <Column Width='150' />
        <Column Width='200' />
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
    <Style Identifier='MediumShading1Accent1' />
</TableDefinition>";

            // Parse the XML.
            XDocument xDoc = XDocument.Parse(tableXml);
            var columnElements = xDoc.Root.Element("Columns")?.Elements("Column") ?? Enumerable.Empty<XElement>();
            var columnWidths = columnElements
                .Select(c => (double?)c.Attribute("Width"))
                .Where(w => w.HasValue)
                .Select(w => w.Value)
                .ToList();

            var rowElements = xDoc.Root.Element("Rows")?.Elements("Row") ?? Enumerable.Empty<XElement>();
            var styleIdentifierStr = (string)xDoc.Root.Element("Style")?.Attribute("Identifier");

            // Create a new document and a builder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start the table.
            Table table = builder.StartTable();

            // Ensure the table has at least one row (required before setting style).
            table.EnsureMinimum();

            // Apply style if it was specified in the XML.
            if (!string.IsNullOrEmpty(styleIdentifierStr) &&
                Enum.TryParse(styleIdentifierStr, out StyleIdentifier styleId))
            {
                table.StyleIdentifier = styleId;
                table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;
            }

            // Use FixedColumnWidths so that the preferred widths we set are respected.
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            // Move the builder cursor into the first cell that EnsureMinimum created.
            builder.MoveTo(table.FirstRow.FirstCell.FirstParagraph);

            int rowIndex = 0;
            foreach (XElement rowElem in rowElements)
            {
                var cellElements = rowElem.Elements("Cell").ToList();

                for (int i = 0; i < cellElements.Count; i++)
                {
                    // For the very first cell (row 0, column 0) we already have a cell created by EnsureMinimum.
                    if (!(rowIndex == 0 && i == 0))
                        builder.InsertCell();

                    // Set the preferred width for the cell if a width was defined for this column.
                    if (i < columnWidths.Count)
                        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(columnWidths[i]);

                    // Write the cell text.
                    builder.Write(cellElements[i].Value);
                }

                // End the current row.
                builder.EndRow();
                rowIndex++;
            }

            // Finish the table.
            builder.EndTable();

            // Define the output path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OutputTable.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file at '{outputPath}'.");
        }
    }
}
