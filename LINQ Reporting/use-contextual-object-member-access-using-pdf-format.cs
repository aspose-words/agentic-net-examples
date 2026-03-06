using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text to the document.
        builder.Writeln("Hello PDF world!");

        // Insert an image so that the document contains a shape.
        // The image will be embedded as a shape node.
        builder.InsertImage(ImageDir + "Sample.png");

        // Save the document to PDF with specific save options.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            // Display the document outline when the PDF is opened.
            PageMode = PdfPageMode.UseOutlines,
            // Ensure that form fields are preserved (optional).
            PreserveFormFields = true
        };
        string pdfPath = Path.Combine(ArtifactsDir, "ContextualObjectAccess.pdf");
        doc.Save(pdfPath, saveOptions);

        // Load the saved PDF with load options that skip images.
        PdfLoadOptions loadOptions = new PdfLoadOptions
        {
            // Skip loading images from the PDF.
            SkipPdfImages = true,
            // Load only the first page (optional).
            PageIndex = 0,
            PageCount = 1
        };
        Document loadedDoc = new Document(pdfPath, loadOptions);

        // Access the collection of shape nodes in the loaded document.
        NodeCollection shapeCollection = loadedDoc.GetChildNodes(NodeType.Shape, true);

        // Output the number of shapes found (should be 0 because images were skipped).
        Console.WriteLine($"Number of shapes in loaded PDF: {shapeCollection.Count}");
    }

    // Helper paths (adjust as needed for your environment).
    private static readonly string ArtifactsDir = @"C:\Output";
    private static readonly string ImageDir = @"C:\Images";
}
