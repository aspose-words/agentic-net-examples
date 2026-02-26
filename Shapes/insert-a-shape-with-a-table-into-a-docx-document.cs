using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using System.Drawing;

// Create a new blank document and associate a DocumentBuilder with it.
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);

// Insert a floating text box shape (inline shapes cannot contain tables).
Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 150);
textBox.WrapType = WrapType.None;                     // No text wrapping.
textBox.HorizontalAlignment = HorizontalAlignment.Center;
textBox.VerticalAlignment = VerticalAlignment.Center;

// Build a simple table that will be placed inside the shape.
Table table = new Table(doc);
table.EnsureMinimum();                               // Guarantees at least one row, cell, paragraph.

// Add some content to the first cell.
Cell firstCell = table.FirstRow.FirstCell;
firstCell.FirstParagraph.AppendChild(new Run(doc, "Hello inside the shape!"));
firstCell.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;

// Insert the table into the text box shape.
textBox.AppendChild(table);

// Save the document to a DOCX file.
doc.Save("ShapeWithTable.docx");
