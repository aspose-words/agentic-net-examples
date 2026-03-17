using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

// Load an existing document that contains at least one table.
Document doc = new Document("Input.docx");

// Retrieve the first table in the document.
Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

// Apply a uniform border thickness of 2 points to all sides of the table.
// The SetBorders method sets the line style, width and colour for every border of the table.
// If the current line style is None, this call automatically changes it to Single.
// Parameters: line style, line width (points), colour.

table.SetBorders(LineStyle.Single, 2.0, Color.Black);

// Save the modified document.

doc.Save("Output.docx");
