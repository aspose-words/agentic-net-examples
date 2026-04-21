using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define paths for the sample source document and the split output folder.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);
        string sourcePath = Path.Combine(dataDir, "Source.docx");
        string outputDir = Path.Combine(dataDir, "SplitParts");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX document with three sections.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        // Section 1
        builder.Writeln("This is the content of Section 1.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2
        builder.Writeln("This is the content of Section 2.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 3
        builder.Writeln("This is the content of Section 3.");

        // Save the sample document to disk.
        sampleDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the DOCX source document using the Document class.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document(sourcePath);

        // -----------------------------------------------------------------
        // 3. Split the document by its sections.
        //    For each section create a new document, import the section,
        //    and save it as a separate DOCX file.
        // -----------------------------------------------------------------
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            // Create a new empty document that will hold the single section.
            Document splitDoc = new Document();

            // Import the section from the source document into the new document.
            // ImportNode performs a deep clone and copies all required styles.
            Section importedSection = (Section)splitDoc.ImportNode(sourceDoc.Sections[i], true);

            // Append the imported section to the new document.
            splitDoc.AppendChild(importedSection);

            // Define the output file name for this split part.
            string partPath = Path.Combine(outputDir, $"Section_{i + 1}.docx");

            // Save the split document.
            splitDoc.Save(partPath);
        }

        // -----------------------------------------------------------------
        // 4. Validate that the expected split output files exist.
        // -----------------------------------------------------------------
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            string partPath = Path.Combine(outputDir, $"Section_{i + 1}.docx");
            if (!File.Exists(partPath))
                throw new FileNotFoundException($"Expected split file not found: {partPath}");
        }

        // The program finishes without requiring any user interaction.
    }
}
