using Aspose.Words;
using Aspose.Words.Saving;

 // Create a new document and add some content.
 Document doc = new Document();
 DocumentBuilder builder = new DocumentBuilder(doc);
 builder.Writeln("Sample text for vertically stretched TIFF.");

 // Configure image save options for TIFF with custom DPI.
 ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
 {
     // Set horizontal DPI (DpiX) to 72.
     HorizontalResolution = 72f,
     // Set vertical DPI (DpiY) to 144 to stretch vertically.
     VerticalResolution = 144f
 };

 // Save the document as a TIFF image using the configured options.
 doc.Save("VerticallyStretched.tiff", tiffOptions);
