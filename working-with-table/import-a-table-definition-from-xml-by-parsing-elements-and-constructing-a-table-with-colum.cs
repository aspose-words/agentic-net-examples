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
            // Sample XML that defines a table structure, column widths and background colors.
            const string xmlContent = @"
<TableDefinition>
    <Columns>
        <Column Index='0' Width='100' Color='LightBlue' />
        <Column Index='1' Width='200' Color='LightGreen' />
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
</TableDefinition>";

            // Parse the XML.
            XDocument xDoc = XDocument.Parse(xmlContent);
            var columnDefs = xDoc.Root
                .Element("Columns")
                .Elements("Column")
                .Select(c => new
                {
                    Index = (int)c.Attribute("Index"),
                    Width = (double)c.Attribute("Width"),
                    ColorName = (string)c.Attribute("Color")
                })
                .OrderBy(c => c.Index)
                .ToList();

            var rowElements = xDoc.Root
                .Element("Rows")
                .Elements("Row")
                .ToList();

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start building the table.
            Table table = builder.StartTable();

            // Iterate over each row defined in the XML.
            foreach (var rowElem in rowElements)
            {
                var cellElements = rowElem.Elements("Cell").ToList();

                for (int i = 0; i < cellElements.Count; i++)
                {
                    // Insert a new cell.
                    builder.InsertCell();

                    // Apply column-specific formatting if a definition exists for this column index.
                    var colDef = columnDefs.FirstOrDefault(c => c.Index == i);
                    if (colDef != null)
                    {
                        // Set the preferred width for the cell.
                        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(colDef.Width);

                        // Set the background shading color.
                        Color bgColor = Color.FromName(colDef.ColorName);
                        if (bgColor.IsKnownColor)
                        {
                            builder.CellFormat.Shading.BackgroundPatternColor = bgColor;
                        }
                    }

                    // Write the cell text.
                    builder.Write(cellElements[i].Value);
                }

                // End the current row.
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Define the output path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OutputTableFromXml.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not created.");

            // The program ends here without waiting for user input.
        }
    }
}
