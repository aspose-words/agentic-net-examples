// Load an existing DOCX file
var doc = new Aspose.Words.Document("InputDocument.docx");

// Create a DocumentBuilder for the loaded document
var builder = new Aspose.Words.DocumentBuilder(doc);

// Write some prompt text before the combo box (optional)
builder.Write("Please select a fruit: ");

// Define the items for the combo box
string[] items = { "Apple", "Banana", "Cherry" };

// Insert the combo box form field at the current cursor position
// Parameters: field name, items array, selected index (0 = first item)
builder.InsertComboBox("FruitComboBox", items, 0);

// Save the modified document
doc.Save("OutputDocument.docx");
