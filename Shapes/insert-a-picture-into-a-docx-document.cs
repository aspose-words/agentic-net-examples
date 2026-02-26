using Aspose.Words;
using Aspose.Words.Drawing;

// Path to the image file to be inserted.
string imagePath = @"C:\Images\Sample.jpg";

// Create a new blank document.
Document doc = new Document();

// Initialize a DocumentBuilder for the document.
DocumentBuilder builder = new DocumentBuilder(doc);

// Insert the image inline at the current cursor position.
builder.InsertImage(imagePath);

// Save the document to a DOCX file.
doc.Save(@"C:\Output\PictureDoc.docx");
