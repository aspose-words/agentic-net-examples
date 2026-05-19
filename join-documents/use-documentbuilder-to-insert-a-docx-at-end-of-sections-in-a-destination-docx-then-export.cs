using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for temporary files.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);

        // -----------------------------------------------------------------
        // 1. Create a source DOCX that will be inserted into the destination.
        // -----------------------------------------------------------------
        string sourcePath = Path.Combine(dataDir, "Source.docx");
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("=== Inserted Document Content ===");
        srcBuilder.Writeln("This text comes from the source DOCX.");
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Create a destination DOCX with several sections.
        // -----------------------------------------------------------------
        string destinationPath = Path.Combine(dataDir, "Destination.docx");
        Document destinationDoc = new Document();
        DocumentBuilder dstBuilder = new DocumentBuilder(destinationDoc);

        for (int i = 1; i <= 3; i++)
        {
            dstBuilder.Writeln($"--- Destination Section {i} ---");
            dstBuilder.Writeln($"Content of section {i}.");
            // Add a section break after each section except the last one.
            if (i < 3)
                dstBuilder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        destinationDoc.Save(destinationPath);

        // -----------------------------------------------------------------
        // 3. Load both documents for the join operation.
        // -----------------------------------------------------------------
        Document dst = new Document(destinationPath);
        Document src = new Document(sourcePath);

        // -----------------------------------------------------------------
        // 4. Insert the source document at the end of each section in the destination.
        // -----------------------------------------------------------------
        DocumentBuilder builder = new DocumentBuilder(dst);

        for (int i = 0; i < dst.Sections.Count; i++)
        {
            // Move the cursor to the last paragraph of the current section.
            Paragraph lastParagraph = dst.Sections[i].Body.LastParagraph;
            builder.MoveTo(lastParagraph);

            // Optional: add a page break before the inserted content for clarity.
            builder.InsertBreak(BreakType.PageBreak);

            // Insert the entire source document at the current cursor position.
            builder.InsertDocument(src, ImportFormatMode.KeepSourceFormatting);

            // If this is not the last section, add another page break to separate from the next section.
            if (i < dst.Sections.Count - 1)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // -----------------------------------------------------------------
        // 5. Save the merged result.
        // -----------------------------------------------------------------
        string mergedPath = Path.Combine(dataDir, "Merged.docx");
        dst.Save(mergedPath);

        // -----------------------------------------------------------------
        // 6. Simple validation to ensure the file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException("The merged document was not saved correctly.");

        // The program finishes here without any interactive prompts.
    }
}
