// Load an existing DOCX file
var doc = new Aspose.Words.Document("Input.docx");

// Create a DocumentBuilder attached to the loaded document
var builder = new Aspose.Words.DocumentBuilder(doc);

// Move the cursor to the first cell of the first table (table index 0, row 0, column 0)
// characterIndex = 0 positions the cursor at the start of the cell
builder.MoveToCell(0, 0, 0, 0);

// Insert the desired text into the cell
builder.Write("Inserted text into the cell");

// Save the modified document
doc.Save("Output.docx");
