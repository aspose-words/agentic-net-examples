using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Math;          // Needed for OfficeMath type
using Aspose.Words.Rendering;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Folder for all generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample document that contains an equation (OfficeMath).
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple equation using the EQ field syntax.
        // The field will be stored as an OfficeMath node.
        builder.InsertField("EQ \\o(\\a\\b,\\c\\d)");

        // Optional: save the sample document for inspection.
        string samplePath = Path.Combine(outputDir, "SampleWithEquation.docx");
        doc.Save(samplePath);

        // ---------------------------------------------------------------
        // 2. Locate all OfficeMath (equation) objects in the document.
        // ---------------------------------------------------------------
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        if (mathNodes.Count == 0)
        {
            Console.WriteLine("No equation objects were found in the document.");
            return;
        }

        // ---------------------------------------------------------------
        // 3. Render each equation to a PNG image and save it.
        // ---------------------------------------------------------------
        int imageIndex = 0;
        foreach (OfficeMath math in mathNodes)
        {
            // Configure image saving options – PNG format with a larger scale.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                Scale = 5 // Render the equation five times larger than its default size.
            };

            string imagePath = Path.Combine(outputDir, $"Equation_{imageIndex}.png");
            math.GetMathRenderer().Save(imagePath, saveOptions);
            Console.WriteLine($"Extracted equation saved as: {imagePath}");
            imageIndex++;
        }

        // ---------------------------------------------------------------
        // 4. Validate that at least one image was produced.
        // ---------------------------------------------------------------
        if (imageIndex == 0)
            throw new InvalidOperationException("Failed to extract any equation images.");
    }
}
