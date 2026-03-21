using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

namespace SmartArtReplacement
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths.
            // Adjust these paths as needed.
            string inputPath = Path.Combine("Input", "DocumentWithSmartArt.docx");
            string outputPath = Path.Combine("Output", "DocumentWithoutSmartArt.docx");

            // Ensure the Input and Output directories exist.
            Directory.CreateDirectory(Path.GetDirectoryName(inputPath)!);
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

            // If the input file does not exist, create a minimal document so the example can run.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                Console.WriteLine("Creating a blank document as a placeholder.");

                var placeholderDoc = new Document();
                var builder = new DocumentBuilder(placeholderDoc);
                builder.Writeln("This is a placeholder document. No SmartArt is present.");
                placeholderDoc.Save(inputPath);
            }

            // Load the document (lifecycle rule: load).
            Document doc = new Document(inputPath);

            // Find all shapes that contain SmartArt.
            var smartArtShapes = doc.GetChildNodes(NodeType.Shape, true)
                                    .Cast<Shape>()
                                    .Where(s => s.HasSmartArt)
                                    .ToList();

            foreach (Shape smartArt in smartArtShapes)
            {
                // Create a simple rectangle shape to act as a placeholder diagram.
                Shape placeholder = new Shape(doc, ShapeType.Rectangle)
                {
                    // Preserve the original shape's geometry.
                    Left = smartArt.Left,
                    Top = smartArt.Top,
                    Width = smartArt.Width,
                    Height = smartArt.Height,
                    // Preserve wrapping and positioning.
                    WrapType = smartArt.WrapType,
                    WrapSide = smartArt.WrapSide,
                    RelativeHorizontalPosition = smartArt.RelativeHorizontalPosition,
                    RelativeVerticalPosition = smartArt.RelativeVerticalPosition,
                    HorizontalAlignment = smartArt.HorizontalAlignment,
                    VerticalAlignment = smartArt.VerticalAlignment,
                    // Optional: give it a visible fill so the replacement is noticeable.
                    FillColor = System.Drawing.Color.LightGray,
                    // Optional: set a label to indicate replacement.
                    AlternativeText = "Replaced SmartArt"
                };

                // Insert the placeholder after the original SmartArt shape.
                smartArt.ParentNode.InsertAfter(placeholder, smartArt);
                // Remove the original SmartArt shape.
                smartArt.Remove();
            }

            // Save the modified document (lifecycle rule: save).
            doc.Save(outputPath);
            Console.WriteLine($"Processing complete. Output saved to: {outputPath}");
        }
    }
}
