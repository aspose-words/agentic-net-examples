using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

// Create a new document and a builder.
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);

// Build a simple 2x2 table.
Table table = builder.StartTable();
builder.InsertCell();
builder.Write("Cell 1");
builder.InsertCell();
builder.Write("Cell 2");
builder.EndRow();
builder.InsertCell();
builder.Write("Cell 3");
builder.InsertCell();
builder.Write("Cell 4");
builder.EndTable();

// Create a custom table style.
TableStyle style = (TableStyle)doc.Styles.Add(StyleType.Table, "TopBottomBorderStyle");

// Configure top border.
style.Borders.Top.LineStyle = LineStyle.Single;
style.Borders.Top.LineWidth = 1.5;
style.Borders.Top.Color = Color.Black;

// Configure bottom border.
style.Borders.Bottom.LineStyle = LineStyle.Single;
style.Borders.Bottom.LineWidth = 1.5;
style.Borders.Bottom.Color = Color.Black;

// Remove left border.
style.Borders.Left.LineStyle = LineStyle.None;
style.Borders.Left.LineWidth = 0;
style.Borders.Left.Color = Color.Empty;

// Remove right border.
style.Borders.Right.LineStyle = LineStyle.None;
style.Borders.Right.LineWidth = 0;
style.Borders.Right.Color = Color.Empty;

// Apply the style to the table.
table.Style = style;

// Save the document.
doc.Save("TopBottomBorderTable.docx");
