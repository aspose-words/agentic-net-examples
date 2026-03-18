using Aspose.Words;
using Aspose.Words.Saving;

// Create a new Word document.
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);

// Add some sample content.
builder.Writeln("Sample low‑resolution TIFF for email attachment.");

// Configure image save options to render the document as a TIFF image
// with a resolution of 72 DPI (both horizontal and vertical).
ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);
tiffOptions.Resolution = 72; // DesiredDpi = 72

// Save the document as a low‑resolution TIFF file.
doc.Save("LowResImage.tiff", tiffOptions);
