using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace OleObjectReplacementExample
{
    public class Program
    {
        public static void Main()
        {
            // Prepare a temporary working folder.
            string tempDir = Path.Combine(Path.GetTempPath(), "OleExample");
            Directory.CreateDirectory(tempDir);

            // Paths for the input document, the image to embed, and the output document.
            string inputDocPath = Path.Combine(tempDir, "Input.docx");
            string newImagePath = Path.Combine(tempDir, "NewImage.png");
            string outputDocPath = Path.Combine(tempDir, "Output.docx");

            // Ensure an image file exists – write a simple red PNG if it does not.
            if (!File.Exists(newImagePath))
            {
                // Base64-encoded 1x1 red PNG image.
                const string redPngBase64 =
                    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==";
                byte[] pngBytes = Convert.FromBase64String(redPngBase64);
                File.WriteAllBytes(newImagePath, pngBytes);
            }

            // If the input document does not exist, create one containing a dummy OLE object.
            if (!File.Exists(inputDocPath))
            {
                // Create a blank document.
                Document doc = new Document();
                DocumentBuilder builder = new DocumentBuilder(doc);

                // Embed a simple text package as an OLE object.
                using (MemoryStream dummyPackage = new MemoryStream(System.Text.Encoding.UTF8.GetBytes("Dummy OLE content")))
                {
                    // "Package" progId creates a generic OLE package.
                    builder.InsertOleObject(dummyPackage, "Package", false, null);
                }

                // Save the document that will later be loaded.
                doc.Save(inputDocPath);
            }

            // Load the existing document that contains the OLE object.
            Document loadedDoc = new Document(inputDocPath);

            // Find the first OLE object shape.
            Shape oldOleShape = (Shape)loadedDoc.GetChild(NodeType.Shape, 0, true);
            if (oldOleShape == null || oldOleShape.ShapeType != ShapeType.OleObject)
                throw new InvalidOperationException("No OLE object found in the document.");

            // Position the builder at the OLE shape.
            DocumentBuilder builderAtOle = new DocumentBuilder(loadedDoc);
            builderAtOle.MoveTo(oldOleShape);

            // Remove the old OLE shape.
            oldOleShape.Remove();

            // Insert the new image as an embedded OLE object (still using the "Package" progId).
            using (FileStream imageStream = File.OpenRead(newImagePath))
            {
                builderAtOle.InsertOleObject(imageStream, "Package", false, null);
            }

            // Save the modified document.
            loadedDoc.Save(outputDocPath);

            Console.WriteLine($"OLE object replaced successfully. Output saved to: {outputDocPath}");
        }
    }
}
