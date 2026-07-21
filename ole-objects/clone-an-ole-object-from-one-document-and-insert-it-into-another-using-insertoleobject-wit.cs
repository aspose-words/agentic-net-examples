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
            // Path where output files will be saved.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // ------------------------------
            // 1. Create a source document and embed an OLE object (a simple text file packaged as OLE).
            // ------------------------------
            Document srcDoc = new Document();
            DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);

            // Prepare some data to embed.
            byte[] sampleData = System.Text.Encoding.UTF8.GetBytes("Hello from embedded OLE object!");
            using (MemoryStream dataStream = new MemoryStream(sampleData))
            {
                // Insert the OLE object as a package. Use "Package" progId.
                Shape oleShape = srcBuilder.InsertOleObject(dataStream, "Package", true, null);
                // Optionally set display name for the package.
                oleShape.OleFormat.OlePackage.FileName = "Sample.txt";
                oleShape.OleFormat.OlePackage.DisplayName = "Sample.txt";
            }

            // Save the source document (optional, just for inspection).
            string srcPath = Path.Combine(outputDir, "Source.docx");
            srcDoc.Save(srcPath);

            // ------------------------------
            // 2. Extract the OLE object's raw data from the source document.
            // ------------------------------
            // Locate the first shape that contains an OLE object.
            Shape srcOleShape = (Shape)srcDoc.GetChild(NodeType.Shape, 0, true);
            OleFormat srcOleFormat = srcOleShape.OleFormat;

            // Retrieve the raw OLE data as a byte array.
            byte[] oleRawData = srcOleFormat.GetRawData();

            // ------------------------------
            // 3. Create a destination document and insert the cloned OLE object using the extracted data.
            // ------------------------------
            Document dstDoc = new Document();
            DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);

            // Insert a paragraph to separate content.
            dstBuilder.Writeln("Cloned OLE object below:");

            // Use the extracted raw data to create a new OLE object in the destination document.
            using (MemoryStream oleDataStream = new MemoryStream(oleRawData))
            {
                // Insert the OLE object. Reuse the original ProgId and insert as normal (not as icon).
                dstBuilder.InsertOleObject(oleDataStream, srcOleFormat.ProgId, false, null);
            }

            // Save the destination document.
            string dstPath = Path.Combine(outputDir, "Destination.docx");
            dstDoc.Save(dstPath);
        }
    }
}
