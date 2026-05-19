using System;
using System.IO;
using System.Xml;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableLayoutExport
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a sample table with custom widths and shading.
            Table table = builder.StartTable();

            // First row
            builder.InsertCell();
            // Set width and background color for the first cell.
            builder.CellFormat.Width = 100; // points
            builder.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.LightBlue;
            builder.Write("Cell 0,0");

            builder.InsertCell();
            builder.CellFormat.Width = 150;
            builder.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.LightGreen;
            builder.Write("Cell 0,1");
            builder.EndRow();

            // Second row
            builder.InsertCell();
            builder.CellFormat.Width = 120;
            builder.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.LightCoral;
            builder.Write("Cell 1,0");

            builder.InsertCell();
            builder.CellFormat.Width = 130;
            builder.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.LightYellow;
            builder.Write("Cell 1,1");
            builder.EndRow();

            builder.EndTable();

            // Save the sample document (optional, demonstrates lifecycle compliance).
            string docPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleTable.docx");
            doc.Save(docPath);

            // Prepare XML writer for the table layout report.
            string xmlPath = Path.Combine(Directory.GetCurrentDirectory(), "TableLayout.xml");
            using (XmlWriter writer = XmlWriter.Create(xmlPath, new XmlWriterSettings { Indent = true }))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("Tables");

                // Iterate through all tables in the document.
                NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
                for (int t = 0; t < tables.Count; t++)
                {
                    Table tbl = (Table)tables[t];
                    writer.WriteStartElement("Table");
                    writer.WriteAttributeString("Index", t.ToString());

                    // Iterate through rows.
                    for (int r = 0; r < tbl.Rows.Count; r++)
                    {
                        Row row = tbl.Rows[r];
                        writer.WriteStartElement("Row");
                        writer.WriteAttributeString("Index", r.ToString());

                        // Iterate through cells.
                        for (int c = 0; c < row.Cells.Count; c++)
                        {
                            Cell cell = row.Cells[c];
                            writer.WriteStartElement("Cell");
                            writer.WriteAttributeString("Index", c.ToString());

                            // Write cell width (points). If not set, default is 0.
                            double width = cell.CellFormat.Width;
                            writer.WriteAttributeString("Width", width.ToString());

                            // Write background color as ARGB hex if set.
                            var bgColor = cell.CellFormat.Shading.BackgroundPatternColor;
                            if (bgColor != System.Drawing.Color.Empty)
                            {
                                string colorHex = $"#{bgColor.A:X2}{bgColor.R:X2}{bgColor.G:X2}{bgColor.B:X2}";
                                writer.WriteAttributeString("BackgroundColor", colorHex);
                            }

                            // Write vertical alignment as string.
                            writer.WriteAttributeString("VerticalAlignment", cell.CellFormat.VerticalAlignment.ToString());

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
}
