using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace OleCloneExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a source document and embed a simple OLE package (a text file) into it.
            Document sourceDoc = new Document();
            DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

            // Prepare dummy data for the OLE package.
            byte[] packageData = System.Text.Encoding.UTF8.GetBytes("Sample OLE package content");
            using (MemoryStream packageStream = new MemoryStream(packageData))
            {
                // Insert the OLE object as an icon. ProgId "Package" denotes a generic OLE package.
                srcBuilder.InsertOleObject(packageStream, "Package", true, null);
            }

            // Save the source document to a file.
            const string sourcePath = "Source.docx";
            sourceDoc.Save(sourcePath);

            // Load the source document (demonstrating loading with default options).
            Document loadedSource = new Document(sourcePath);

            // Locate the first shape that contains an OLE object.
            Shape sourceOleShape = (Shape)loadedSource.GetChild(NodeType.Shape, 0, true);
            OleFormat sourceOleFormat = sourceOleShape.OleFormat;

            // Extract the embedded OLE data into a memory stream.
            using (MemoryStream extractedOleData = new MemoryStream())
            {
                sourceOleFormat.Save(extractedOleData);
                extractedOleData.Position = 0; // Reset stream position for reading.

                // Create a destination document.
                Document destDoc = new Document();
                DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

                // Insert the extracted OLE data into the destination document.
                // Use the original ProgId and insert as normal (not as an icon).
                destBuilder.InsertOleObject(extractedOleData, sourceOleFormat.ProgId, false, null);

                // Save the destination document.
                const string destPath = "Destination.docx";
                destDoc.Save(destPath);
            }
        }
    }
}
