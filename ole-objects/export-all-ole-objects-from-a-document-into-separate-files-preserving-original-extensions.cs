using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class ExportOleObjects
{
    public static void Main()
    {
        // Path to the source Word document.
        const string inputPath = "Input.docx";

        // Directory where extracted OLE objects will be saved.
        const string outputDir = "ExtractedOleObjects";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputDir);

        // If the input document does not exist, create a sample document that contains an OLE object.
        if (!File.Exists(inputPath))
        {
            // Create a simple text file to embed as an OLE package.
            const string tempTextFile = "Sample.txt";
            File.WriteAllText(tempTextFile, "This is a sample embedded file.");

            // Create a new blank document.
            Document sampleDoc = new Document();

            // Insert the text file as an OLE package (embedded, not linked).
            using (FileStream stream = File.OpenRead(tempTextFile))
            {
                // InsertOleObject(stream, progId, asIcon, presentation)
                // progId "Package" indicates a generic OLE package.
                // asIcon = false to display the content.
                Shape oleShape = new DocumentBuilder(sampleDoc).InsertOleObject(stream, "Package", false, null);
                // Optionally set a display name for the package.
                oleShape.OleFormat.OlePackage.FileName = tempTextFile;
                oleShape.OleFormat.OlePackage.DisplayName = "Sample.txt";
            }

            // Save the sample document.
            sampleDoc.Save(inputPath);
        }

        // Load the document.
        Document doc = new Document(inputPath);

        // Get all shape nodes in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        int oleIndex = 0;

        foreach (Shape shape in shapes)
        {
            // Access the OLE format of the shape, if any.
            OleFormat ole = shape.OleFormat;
            if (ole == null)
                continue;

            // Skip linked OLE objects because saving them throws an exception.
            if (ole.IsLink)
                continue;

            // Determine a file name for the extracted object.
            string fileName = ole.SuggestedFileName;
            if (string.IsNullOrEmpty(fileName))
            {
                // Fallback to a generated name using the suggested extension.
                string extension = ole.SuggestedExtension ?? ".bin";
                fileName = $"OleObject_{oleIndex}{extension}";
            }

            // Combine the output directory with the file name.
            string outputPath = Path.Combine(outputDir, fileName);

            // Save the OLE object to the file system.
            ole.Save(outputPath);

            oleIndex++;
        }

        Console.WriteLine($"Extracted {oleIndex} OLE object(s) to folder '{outputDir}'.");
    }
}
