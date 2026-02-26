// Create a new DOCM document
var doc = new Aspose.Words.Document();

// Initialize a DocumentBuilder for the document
var builder = new Aspose.Words.DocumentBuilder(doc);

// Insert a text input form field at the current cursor position
// Parameters: name, type, format, default text, max length (0 = unlimited)
builder.InsertTextInput(
    name: "MyTextInput",
    type: Aspose.Words.Fields.TextFormFieldType.Regular,
    format: "",
    fieldValue: "Enter your text here",
    maxLength: 0);

// Save the document as a macro‑enabled Word file
doc.Save("FormWithTextInput.docm");
