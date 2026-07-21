using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitDocumentExample
{
    public static void Main()
    {
        // Define an output directory and ensure it exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Path for the sample source DOCX file.
        string sourceDocPath = Path.Combine(outputDir, "Sample.docx");

        // Create a sample document with three sections.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Section 1 - First part");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2 - Second part");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 3 - Third part");
        sampleDoc.Save(sourceDocPath);

        // Load the DOCX source document using the Document class.
        Document doc = new Document(sourceDocPath);

        // Configure HtmlSaveOptions to split the document at each section break.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak
        };

        // Save the document; Aspose.Words will create multiple HTML files.
        string baseHtmlPath = Path.Combine(outputDir, "SplitDocument.html");
        doc.Save(baseHtmlPath, saveOptions);

        // Validate that the split output files were created.
        bool baseExists = File.Exists(baseHtmlPath);
        bool part1Exists = File.Exists(Path.Combine(outputDir, "SplitDocument-01.html"));
        bool part2Exists = File.Exists(Path.Combine(outputDir, "SplitDocument-02.html"));

        if (!baseExists || !part1Exists || !part2Exists)
        {
            throw new Exception("One or more split HTML files were not created as expected.");
        }

        // Indicate successful completion.
        Console.WriteLine("Document split completed successfully. Files are located in: " + outputDir);
    }
}
