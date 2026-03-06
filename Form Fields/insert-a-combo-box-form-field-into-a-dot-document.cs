// Create a new blank Word document (template)
var doc = new Aspose.Words.Document();

// Initialize a DocumentBuilder for the document
var builder = new Aspose.Words.DocumentBuilder(doc);

// Write some prompt text before the combo box
builder.Write("Pick a fruit: ");

// Define the items that will appear in the combo box
string[] items = { "Apple", "Banana", "Cherry" };

// Insert the combo box form field at the current cursor position
// Parameters: field name, array of items, index of the initially selected item
builder.InsertComboBox("FruitComboBox", items, 0);

// Save the document as a DOT (Word template) file
doc.Save("ComboBoxTemplate.dot", Aspose.Words.SaveFormat.Dot);
