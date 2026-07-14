using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace ShapeAlternativeTextExample
{
    public class Program
    {
        public static void Main()
        {
            // Define output directory and file name.
            string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
            Directory.CreateDirectory(artifactsDir);
            string outputPath = Path.Combine(artifactsDir, "Shape.AltText.docx");

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a cube shape and give it a name.
            Shape shape = builder.InsertShape(ShapeType.Cube, 150, 150);
            shape.Name = "MyCube";

            // Set the alternative text for accessibility.
            string altText = "Alt text for MyCube.";
            shape.AlternativeText = altText;

            // Save the document.
            doc.Save(outputPath);

            // Validate that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception($"Failed to create the output file at '{outputPath}'.");

            // Reload the document and verify the alternative text.
            Document loadedDoc = new Document(outputPath);
            Shape loadedShape = (Shape)loadedDoc.GetChild(NodeType.Shape, 0, true);
            if (loadedShape == null)
                throw new Exception("No shape was found in the saved document.");

            if (loadedShape.AlternativeText != altText)
                throw new Exception("The alternative text of the shape does not match the expected value.");

            // Indicate success (optional).
            Console.WriteLine("Shape alternative text set and verified successfully.");
        }
    }
}
