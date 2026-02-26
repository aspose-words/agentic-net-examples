// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.SmartArt;

class InsertSmartArtExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an inline SmartArt shape with the desired size (width, height in points).
        Shape smartArtShape = builder.InsertShape(ShapeType.SmartArt, 400, 300);

        // Set the SmartArt layout (e.g., Basic Cycle).
        smartArtShape.SmartArt.Layout = SmartArtLayout.BasicCycle;

        // Access the root node of the SmartArt diagram.
        SmartArtNode rootNode = smartArtShape.SmartArt.Nodes[0];
        rootNode.TextFrame.Text = "Start";

        // Add a child node to the root.
        SmartArtNode middleNode = rootNode.AddNode();
        middleNode.TextFrame.Text = "Middle";

        // Add a sub‑child node.
        SmartArtNode endNode = middleNode.AddNode();
        endNode.TextFrame.Text = "End";

        // Render the SmartArt drawing (necessary if the pre‑rendered drawing is missing or outdated).
        smartArtShape.UpdateSmartArtDrawing();

        // Save the document to a DOCX file.
        doc.Save("SmartArtDiagram.docx");
    }
}
