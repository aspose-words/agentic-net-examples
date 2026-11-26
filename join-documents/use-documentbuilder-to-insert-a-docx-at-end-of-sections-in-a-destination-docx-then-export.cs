using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the sample documents.
        string destinationPath = "Destination.docx";
        string sourcePath = "Source.docx";
        string mergedPath = "Merged.docx";

        // -----------------------------------------------------------------
        // Create a destination document with two sections.
        // -----------------------------------------------------------------
        Document destinationDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destinationDoc);
        destBuilder.Writeln("Section 1 - Original content.");
        destBuilder.InsertBreak(BreakType.SectionBreakNewPage);
        destBuilder.Writeln("Section 2 - Original content.");
        destinationDoc.Save(destinationPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create a source document that will be inserted.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("Inserted document content.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Load the documents for processing.
        // -----------------------------------------------------------------
        Document dest = new Document(destinationPath);
        Document src = new Document(sourcePath);

        // -----------------------------------------------------------------
        // Insert the source document at the end of each original section.
        // -----------------------------------------------------------------
        int originalSectionCount = dest.Sections.Count; // Preserve original count.
        for (int i = 0; i < originalSectionCount; i++)
        {
            // Move the builder to the end of the current section.
            DocumentBuilder insertBuilder = new DocumentBuilder(dest);
            insertBuilder.MoveToSection(i);
            // Ensure we are positioned after the last node of the section.
            insertBuilder.MoveTo(insertBuilder.CurrentParagraph);
            // Insert a section break to separate the inserted content.
            insertBuilder.InsertBreak(BreakType.SectionBreakNewPage);
            // Insert the whole source document after the break.
            insertBuilder.InsertDocument(src, ImportFormatMode.KeepSourceFormatting);
        }

        // -----------------------------------------------------------------
        // Save the merged document.
        // -----------------------------------------------------------------
        dest.Save(mergedPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Validation: ensure the merged file exists and contains inserted content.
        // -----------------------------------------------------------------
        if (!File.Exists(mergedPath))
        {
            throw new Exception("Merged document was not created.");
        }

        Document merged = new Document(mergedPath);
        // Expect more sections than the original two because each insertion adds a section.
        if (merged.Sections.Count <= 2)
        {
            throw new Exception("Merged document does not contain the inserted sections.");
        }

        // Simple content check.
        if (!merged.GetText().Contains("Inserted document content."))
        {
            throw new Exception("Inserted content was not found in the merged document.");
        }
    }
}
