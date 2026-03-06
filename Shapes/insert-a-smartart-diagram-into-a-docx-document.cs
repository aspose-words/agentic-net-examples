using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertSmartArtExample
{
    static void Main()
    {
        // Path to a template document that already contains a SmartArt shape.
        // The template can be created manually in Microsoft Word and saved as "SmartArtTemplate.docx".
        string templatePath = "SmartArtTemplate.docx";

        // Path where the resulting document will be saved.
        string outputPath = "SmartArtDocument.docx";

        // Load the template document.
        Document templateDoc = new Document(templatePath);

        // Find the first SmartArt shape in the template.
        // SmartArt objects are stored as Shape nodes.
        Shape smartArtShape = (Shape)templateDoc.GetChild(NodeType.Shape, 0, true);

        // Ensure the SmartArt drawing is rendered correctly.
        smartArtShape.UpdateSmartArtDrawing();

        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder for the new document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Import the SmartArt shape into the new document's node collection.
        Shape importedSmartArt = (Shape)doc.ImportNode(smartArtShape, true);

        // Insert the imported SmartArt shape at the current cursor position.
        builder.InsertNode(importedSmartArt);

        // Save the document containing the SmartArt diagram.
        doc.Save(outputPath);
    }
}
