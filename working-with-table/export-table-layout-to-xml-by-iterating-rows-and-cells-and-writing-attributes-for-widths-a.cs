using System;
using System.IO;
using System.Xml;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to construct a sample table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table with custom widths and shading.
        builder.StartTable();

        // First cell
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
        builder.CellFormat.Shading.ForegroundPatternColor = System.Drawing.Color.LightBlue;
        builder.Writeln("Cell 1,1");

        // Second cell
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(150);
        builder.CellFormat.Shading.ForegroundPatternColor = System.Drawing.Color.LightGreen;
        builder.Writeln("Cell 1,2");
        builder.EndRow();

        // Second row, first cell
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(120);
        builder.CellFormat.Shading.ForegroundPatternColor = System.Drawing.Color.LightCoral;
        builder.Writeln("Cell 2,1");

        // Second row, second cell
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(130);
        builder.CellFormat.Shading.ForegroundPatternColor = System.Drawing.Color.LightYellow;
        builder.Writeln("Cell 2,2");
        builder.EndRow();

        builder.EndTable();

        // Save the document (optional, just to have a physical file).
        string docPath = "SampleTable.docx";
        doc.Save(docPath);

        // Prepare XML writer for exporting table layout.
        string xmlPath = "TableLayout.xml";
        XmlWriterSettings settings = new XmlWriterSettings
        {
            Indent = true,
            Encoding = System.Text.Encoding.UTF8
        };

        using (XmlWriter writer = XmlWriter.Create(xmlPath, settings))
        {
            writer.WriteStartDocument();
            writer.WriteStartElement("Tables");

            // Iterate through all tables in the document.
            NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
            int tableIndex = 0;
            foreach (Table table in tables)
            {
                writer.WriteStartElement("Table");
                writer.WriteAttributeString("Index", tableIndex.ToString());

                // Iterate rows.
                for (int rowIdx = 0; rowIdx < table.Rows.Count; rowIdx++)
                {
                    Row row = table.Rows[rowIdx];
                    writer.WriteStartElement("Row");
                    writer.WriteAttributeString("Index", rowIdx.ToString());

                    // Iterate cells.
                    for (int cellIdx = 0; cellIdx < row.Cells.Count; cellIdx++)
                    {
                        Cell cell = row.Cells[cellIdx];
                        writer.WriteStartElement("Cell");
                        writer.WriteAttributeString("RowIndex", rowIdx.ToString());
                        writer.WriteAttributeString("ColumnIndex", cellIdx.ToString());

                        // Width (points) – use PreferredWidth if set.
                        double width = 0;
                        if (cell.CellFormat.PreferredWidth != null &&
                            cell.CellFormat.PreferredWidth.Type == PreferredWidthType.Points)
                        {
                            width = cell.CellFormat.PreferredWidth.Value;
                        }
                        writer.WriteAttributeString("WidthPoints", width.ToString());

                        // Shading color (if any).
                        var shadingColor = cell.CellFormat.Shading.ForegroundPatternColor;
                        writer.WriteAttributeString("ShadingColor", shadingColor.IsEmpty ? "None" : shadingColor.Name);

                        writer.WriteEndElement(); // Cell
                    }

                    writer.WriteEndElement(); // Row
                }

                writer.WriteEndElement(); // Table
                tableIndex++;
            }

            writer.WriteEndElement(); // Tables
            writer.WriteEndDocument();
        }

        // Validate that the XML file was created.
        if (!File.Exists(xmlPath))
        {
            throw new Exception($"Failed to create XML report at '{xmlPath}'.");
        }

        // Optionally, output a simple confirmation (no user interaction required).
        Console.WriteLine("Table layout exported to XML successfully.");
    }
}
