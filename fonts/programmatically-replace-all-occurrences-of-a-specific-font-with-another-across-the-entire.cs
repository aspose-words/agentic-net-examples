using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Prepare directories and file paths.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);
        string inputPath = Path.Combine(dataDir, "Sample.docx");
        string outputPath = Path.Combine(dataDir, "Result.docx");

        // Create a sample document containing the fonts to be replaced.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Font.Name = "Arial";
        builder.Writeln("This paragraph uses Arial.");
        builder.Font.Name = "Times New Roman";
        builder.Writeln("This paragraph uses Times New Roman.");
        builder.Font.Name = "Arial";
        builder.Writeln("Another paragraph with Arial.");

        // Save the sample document.
        sampleDoc.Save(inputPath);

        // Load the document for processing.
        Document doc = new Document(inputPath);

        // Define the source font to replace and the target font.
        const string sourceFont = "Arial";
        const string targetFont = "Courier New";

        // Iterate over all Run nodes and replace the font where it matches the source font.
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            if (run.Font.Name.Equals(sourceFont, StringComparison.OrdinalIgnoreCase))
                run.Font.Name = targetFont;
        }

        // Validate that no Run still uses the source font.
        bool replacementSuccessful = doc.GetChildNodes(NodeType.Run, true)
            .Cast<Run>()
            .All(r => !r.Font.Name.Equals(sourceFont, StringComparison.OrdinalIgnoreCase));

        // Save the modified document.
        doc.Save(outputPath);

        // Ensure the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");

        // Output validation result.
        Console.WriteLine(replacementSuccessful
            ? "All occurrences of the source font were successfully replaced."
            : "Some occurrences of the source font remain.");
    }
}
