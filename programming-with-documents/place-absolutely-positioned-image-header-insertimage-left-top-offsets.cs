using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder's cursor to the primary header of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Insert a floating image into the header.
        // The image is positioned relative to the page (absolute positioning).
        // Left offset = 50 points, top offset = 30 points.
        // Width and height are set to -1 to keep the original image dimensions.
        // WrapType.None places the image behind the text.
        Shape shape = builder.InsertImage(
            "ImageDir/Logo.png",                     // Path to the image file.
            RelativeHorizontalPosition.Page,         // Horizontal reference (page).
            50,                                      // Left offset (points).
            RelativeVerticalPosition.Page,           // Vertical reference (page).
            30,                                      // Top offset (points).
            -1,                                      // Width (auto).
            -1,                                      // Height (auto).
            WrapType.None);                          // No text wrapping.

        // Ensure the image appears behind any overlapping text.
        shape.BehindText = true;

        // Save the resulting document.
        doc.Save("Output/HeaderImage.docx");
    }
}
