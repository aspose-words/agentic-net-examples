// Load an existing DOTM template
var doc = new Aspose.Words.Document("Template.dotm");

// Create a DocumentBuilder attached to the loaded document
var builder = new Aspose.Words.DocumentBuilder(doc);

// Move the cursor to the desired cell.
 // Parameters: tableIndex, rowIndex, columnIndex, cellIndex
 // Adjust the indices as needed for your document.
builder.MoveToCell(0, 0, 0, 0);

// Insert the text into the cell
builder.Write("Inserted text");

// Save the modified document
doc.Save("Result.dotm");
