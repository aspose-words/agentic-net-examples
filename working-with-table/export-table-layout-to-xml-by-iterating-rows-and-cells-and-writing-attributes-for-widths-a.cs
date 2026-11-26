using System;
using System.IO;
using System.Drawing;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableLayoutExport
{
    class Program
    {
        static void Main()
        {
            // Create a new document and a builder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple table with some formatting.
            builder.StartTable();

            // First row
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(80);
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
            builder.Writeln("Cell 1-1");

            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(120);
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
            builder.Writeln("Cell 1-2");

            builder.EndRow();

            // Second row
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightCoral;
            builder.Writeln("Cell 2-1");

            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightYellow;
            builder.Writeln("Cell 2-2");

            builder.EndRow();

            builder.EndTable();

            // Save the sample document.
            string docPath = "SampleTable.docx";
            doc.Save(docPath);

            // Prepare XML document for layout export.
            XDocument xmlDoc = new XDocument(new XElement("TableLayout"));

            // Iterate through all tables in the document.
            NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
            int tableIndex = 0;
            foreach (Table table in tables)
            {
                XElement tableElement = new XElement("Table",
                    new XAttribute("Index", tableIndex));

                int rowIndex = 0;
                foreach (Row row in table.Rows)
                {
                    XElement rowElement = new XElement("Row",
                        new XAttribute("Index", rowIndex));

                    int cellIndex = 0;
                    foreach (Cell cell in row.Cells)
                    {
                        // Get cell width.
                        double width = 0;
                        if (cell.CellFormat.PreferredWidth != null)
                            width = cell.CellFormat.PreferredWidth.Value;

                        // Get cell shading color.
                        Color shadingColor = cell.CellFormat.Shading?.BackgroundPatternColor ?? Color.Empty;
                        string colorHex = shadingColor.IsEmpty ? "" : $"#{shadingColor.R:X2}{shadingColor.G:X2}{shadingColor.B:X2}";

                        XElement cellElement = new XElement("Cell",
                            new XAttribute("Index", cellIndex),
                            new XAttribute("WidthPoints", width),
                            new XAttribute("ShadingColor", colorHex));

                        rowElement.Add(cellElement);
                        cellIndex++;
                    }

                    tableElement.Add(rowElement);
                    rowIndex++;
                }

                xmlDoc.Root.Add(tableElement);
                tableIndex++;
            }

            // Save the XML layout file.
            string xmlPath = "TableLayout.xml";
            xmlDoc.Save(xmlPath);

            // Simple validation that files were created.
            if (!File.Exists(docPath))
                throw new Exception($"Document file was not created: {docPath}");
            if (!File.Exists(xmlPath))
                throw new Exception($"XML layout file was not created: {xmlPath}");
        }
    }
}
