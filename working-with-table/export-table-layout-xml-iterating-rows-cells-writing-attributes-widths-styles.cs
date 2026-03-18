using System;
using System.IO;
using System.Xml;
using System.Drawing; // Added for Color handling
using Aspose.Words;
using Aspose.Words.Tables;

class TableLayoutExporter
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("Input.docx");

        // Prepare XML writer settings.
        XmlWriterSettings xmlSettings = new XmlWriterSettings
        {
            Indent = true,
            IndentChars = "  ",
            NewLineOnAttributes = false
        };

        // Export table layout to XML.
        using (XmlWriter writer = XmlWriter.Create("TableLayout.xml", xmlSettings))
        {
            writer.WriteStartDocument();
            writer.WriteStartElement("Tables");

            // Get all tables in the document (including nested tables).
            NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);

            for (int t = 0; t < tables.Count; t++)
            {
                Table table = (Table)tables[t];
                writer.WriteStartElement("Table");
                writer.WriteAttributeString("Index", t.ToString());

                // Table level properties (example: cell spacing, alignment).
                writer.WriteAttributeString("CellSpacing", table.CellSpacing.ToString());
                writer.WriteAttributeString("Alignment", table.Alignment.ToString());

                // Iterate rows.
                for (int r = 0; r < table.Rows.Count; r++)
                {
                    Row row = table.Rows[r];
                    writer.WriteStartElement("Row");
                    writer.WriteAttributeString("Index", r.ToString());

                    // Row level properties (example: height, height rule).
                    writer.WriteAttributeString("Height", row.RowFormat.Height.ToString());
                    writer.WriteAttributeString("HeightRule", row.RowFormat.HeightRule.ToString());

                    // Iterate cells.
                    for (int c = 0; c < row.Cells.Count; c++)
                    {
                        Cell cell = row.Cells[c];
                        writer.WriteStartElement("Cell");
                        writer.WriteAttributeString("Index", c.ToString());

                        // Cell width (points). If width is not set, the value is 0.
                        writer.WriteAttributeString("Width", cell.CellFormat.Width.ToString());

                        // Vertical alignment of the cell content.
                        writer.WriteAttributeString("VerticalAlignment", cell.CellFormat.VerticalAlignment.ToString());

                        // Background color (if any) – converted to HTML hex string.
                        Color bgColor = cell.CellFormat.Shading?.BackgroundPatternColor ?? Color.Empty;
                        if (bgColor != Color.Empty)
                        {
                            string htmlColor = $"#{bgColor.R:X2}{bgColor.G:X2}{bgColor.B:X2}";
                            writer.WriteAttributeString("BackgroundColor", htmlColor);
                        }

                        // Write the plain text content of the cell.
                        string cellText = cell.ToString(SaveFormat.Text).Trim();
                        writer.WriteElementString("Text", cellText);

                        writer.WriteEndElement(); // Cell
                    }

                    writer.WriteEndElement(); // Row
                }

                writer.WriteEndElement(); // Table
            }

            writer.WriteEndElement(); // Tables
            writer.WriteEndDocument();
        }

        Console.WriteLine("Table layout exported to TableLayout.xml");
    }
}
