using System;
using System.IO;
using System.Xml;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table with different cell widths and shading.
        builder.StartTable();

        // First row, first cell.
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(80);
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
        builder.InsertCell();
        builder.Writeln("R1C1");

        // First row, second cell.
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(120);
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
        builder.InsertCell();
        builder.Writeln("R1C2");

        builder.EndRow();

        // Second row, first cell.
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightCoral;
        builder.InsertCell();
        builder.Writeln("R2C1");

        // Second row, second cell.
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightYellow;
        builder.InsertCell();
        builder.Writeln("R2C2");

        builder.EndRow();
        builder.EndTable();

        // Save the document (optional, just to have a physical file).
        string docPath = "Sample.docx";
        doc.Save(docPath);

        // Prepare XML writer for the layout export.
        string xmlPath = "TableLayout.xml";
        XmlWriterSettings settings = new XmlWriterSettings { Indent = true };
        using (XmlWriter writer = XmlWriter.Create(xmlPath, settings))
        {
            writer.WriteStartDocument();
            writer.WriteStartElement("Tables");

            // Iterate through all tables in the document.
            NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
            for (int t = 0; t < tables.Count; t++)
            {
                Table table = (Table)tables[t];
                writer.WriteStartElement("Table");
                writer.WriteAttributeString("Index", t.ToString());

                // Iterate rows.
                for (int r = 0; r < table.Rows.Count; r++)
                {
                    Row row = table.Rows[r];
                    writer.WriteStartElement("Row");
                    writer.WriteAttributeString("Index", r.ToString());

                    // Iterate cells.
                    for (int c = 0; c < row.Cells.Count; c++)
                    {
                        Cell cell = row.Cells[c];
                        writer.WriteStartElement("Cell");
                        writer.WriteAttributeString("Index", c.ToString());

                        // Width (points). If not set, default to 0.
                        double width = 0;
                        PreferredWidth prefWidth = cell.CellFormat.PreferredWidth;
                        if (prefWidth != null && prefWidth.Type == PreferredWidthType.Points)
                            width = prefWidth.Value;
                        writer.WriteAttributeString("Width", width.ToString());

                        // Shading color as hex string.
                        Color shade = cell.CellFormat.Shading.BackgroundPatternColor;
                        string shadeHex = ColorTranslator.ToHtml(shade);
                        writer.WriteAttributeString("ShadingColor", shadeHex);

                        // Cell text content (optional, for completeness).
                        writer.WriteString(cell.GetText().TrimEnd('\a'));

                        writer.WriteEndElement(); // Cell
                    }

                    writer.WriteEndElement(); // Row
                }

                writer.WriteEndElement(); // Table
            }

            writer.WriteEndElement(); // Tables
            writer.WriteEndDocument();
        }

        // Verify that the XML file was created.
        if (!File.Exists(xmlPath))
            throw new FileNotFoundException("Failed to create the XML layout file.", xmlPath);
    }
}
