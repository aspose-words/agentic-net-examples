using System;
using System.IO;
using System.Xml;
using System.Drawing; // Needed for ColorTranslator
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

namespace TableLayoutExport
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple table with formatting.
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.CellFormat.Width = 100; // Set cell width.
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
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
            builder.CellFormat.Shading.BackgroundPatternColor = Color.White;
            builder.Write("Row2 Col1");

            builder.InsertCell();
            builder.CellFormat.Width = 150;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.White;
            builder.Write("Row2 Col2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Save the document to a local file.
            string docPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleTable.docx");
            doc.Save(docPath);

            // Export table layout to XML.
            string xmlPath = Path.Combine(Directory.GetCurrentDirectory(), "TableLayout.xml");
            ExportTableLayoutToXml(doc, xmlPath);

            // Simple validation that files were created.
            if (!File.Exists(docPath))
                throw new FileNotFoundException("Document file was not created.", docPath);
            if (!File.Exists(xmlPath))
                throw new FileNotFoundException("XML layout file was not created.", xmlPath);
        }

        private static void ExportTableLayoutToXml(Document doc, string xmlFilePath)
        {
            // Prepare an XML writer with indentation for readability.
            XmlWriterSettings settings = new XmlWriterSettings
            {
                Indent = true,
                IndentChars = "  "
            };

            using (XmlWriter writer = XmlWriter.Create(xmlFilePath, settings))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("Tables");

                // Get all tables in the document.
                NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
                for (int t = 0; t < tables.Count; t++)
                {
                    Table table = (Table)tables[t];
                    writer.WriteStartElement("Table");
                    writer.WriteAttributeString("Index", t.ToString());

                    // Iterate through rows.
                    for (int r = 0; r < table.Rows.Count; r++)
                    {
                        Row row = table.Rows[r];
                        writer.WriteStartElement("Row");
                        writer.WriteAttributeString("Index", r.ToString());

                        // Iterate through cells.
                        for (int c = 0; c < row.Cells.Count; c++)
                        {
                            Cell cell = row.Cells[c];
                            writer.WriteStartElement("Cell");
                            writer.WriteAttributeString("Index", c.ToString());

                            // Write cell width.
                            double width = cell.CellFormat.Width;
                            writer.WriteAttributeString("Width", width.ToString());

                            // Write shading (background) color if set.
                            Color bgColor = cell.CellFormat.Shading.BackgroundPatternColor;
                            if (bgColor != Color.Empty && bgColor != Color.Transparent)
                            {
                                string hex = ColorTranslator.ToHtml(bgColor);
                                writer.WriteAttributeString("BackgroundColor", hex);
                            }

                            // Write the plain text content of the cell.
                            string cellText = cell.ToString(SaveFormat.Text).Trim();
                            writer.WriteString(cellText);

                            writer.WriteEndElement(); // Cell
                        }

                        writer.WriteEndElement(); // Row
                    }

                    writer.WriteEndElement(); // Table
                }

                writer.WriteEndElement(); // Tables
                writer.WriteEndDocument();
            }
        }
    }
}
