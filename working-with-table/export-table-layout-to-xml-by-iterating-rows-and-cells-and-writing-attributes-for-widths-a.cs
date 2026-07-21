using System;
using System.IO;
using System.Xml;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;
using System.Drawing;

public class ExportTableLayout
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a sample table with formatting.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.CellFormat.Width = 100; // Set cell width.
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue; // Set shading.
        builder.Write("Header 1");

        builder.InsertCell();
        builder.CellFormat.Width = 150;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
        builder.Write("Header 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.CellFormat.Width = 100;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.White;
        builder.Write("Row1 Col1");

        builder.InsertCell();
        builder.CellFormat.Width = 150;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.White;
        builder.Write("Row1 Col2");
        builder.EndRow();

        // Third row.
        builder.InsertCell();
        builder.CellFormat.Width = 100;
        builder.Write("Row2 Col1");

        builder.InsertCell();
        builder.CellFormat.Width = 150;
        builder.Write("Row2 Col2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Define output paths.
        string docPath = Path.Combine(Directory.GetCurrentDirectory(), "TableSample.docx");
        string xmlPath = Path.Combine(Directory.GetCurrentDirectory(), "TableLayout.xml");

        // Save the document.
        doc.Save(docPath);

        // Export table layout to XML.
        using (XmlWriter writer = XmlWriter.Create(xmlPath, new XmlWriterSettings { Indent = true }))
        {
            writer.WriteStartDocument();
            writer.WriteStartElement("Tables");

            // Get all tables in the document.
            NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
            for (int t = 0; t < tables.Count; t++)
            {
                Table tbl = (Table)tables[t];
                writer.WriteStartElement("Table");
                writer.WriteAttributeString("Index", t.ToString());

                // Iterate rows.
                for (int r = 0; r < tbl.Rows.Count; r++)
                {
                    Row row = tbl.Rows[r];
                    writer.WriteStartElement("Row");
                    writer.WriteAttributeString("Index", r.ToString());

                    // Iterate cells.
                    for (int c = 0; c < row.Cells.Count; c++)
                    {
                        Cell cell = row.Cells[c];
                        writer.WriteStartElement("Cell");
                        writer.WriteAttributeString("Index", c.ToString());

                        // Write cell width if set.
                        double width = cell.CellFormat.Width;
                        if (width > 0)
                            writer.WriteAttributeString("Width", width.ToString());

                        // Write shading color if not empty.
                        Color shading = cell.CellFormat.Shading.BackgroundPatternColor;
                        if (shading != Color.Empty && shading != Color.Transparent)
                            writer.WriteAttributeString("ShadingColor", ColorTranslator.ToHtml(shading));

                        // Write cell text content.
                        string text = cell.ToString(SaveFormat.Text).Trim();
                        writer.WriteElementString("Text", text);

                        writer.WriteEndElement(); // Cell
                    }

                    writer.WriteEndElement(); // Row
                }

                writer.WriteEndElement(); // Table
            }

            writer.WriteEndElement(); // Tables
            writer.WriteEndDocument();
        }

        // Verify that both files were created.
        if (!File.Exists(docPath))
            throw new FileNotFoundException("Document file was not created.", docPath);
        if (!File.Exists(xmlPath))
            throw new FileNotFoundException("XML layout file was not created.", xmlPath);
    }
}
