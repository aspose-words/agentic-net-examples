using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;

public class Program
{
    public static void Main()
    {
        // Create a sample destination DOCX with two sections.
        var destinationPath = "Destination.docx";
        var destDoc = new Document();
        var destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Destination Section 1");
        destBuilder.InsertBreak(BreakType.SectionBreakNewPage);
        destBuilder.Writeln("Destination Section 2");
        destDoc.Save(destinationPath, SaveFormat.Docx);

        // Create a sample source DOCX to be inserted.
        var sourcePath = "Source.docx";
        var sourceDoc = new Document();
        var sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("Source Content");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // Load the documents for joining.
        var destination = new Document(destinationPath);
        var source = new Document(sourcePath);

        // Remember the original number of sections in the destination document.
        int originalSectionCount = destination.Sections.Count;

        // Use DocumentBuilder to insert the source document at the end of each original section.
        var builder = new DocumentBuilder(destination);
        for (int i = 0; i < originalSectionCount; i++)
        {
            Section section = destination.Sections[i];
            // Move the builder to the last paragraph of the current section.
            builder.MoveTo(section.Body.LastParagraph);
            // Insert the entire source document after the current position.
            builder.InsertDocument(source, ImportFormatMode.KeepSourceFormatting);
        }

        // Save the merged document.
        var mergedPath = "MergedOutput.docx";
        destination.Save(mergedPath, SaveFormat.Docx);

        // Validate that the merged file was created.
        if (!File.Exists(mergedPath))
        {
            throw new InvalidOperationException("Merged output file was not created.");
        }

        // Optional: indicate success (no interactive input required).
        Console.WriteLine("Documents merged successfully.");
    }
}
