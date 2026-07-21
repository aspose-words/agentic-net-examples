using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with two sections.
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Content of Section 1");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Content of Section 2");
        sourceDoc.Save(sourcePath);

        // Load the document to be split.
        Document doc = new Document(sourcePath);

        // Create a DocumentSplitCriteria instance and set it to split by sections.
        DocumentSplitCriteria splitCriteria = DocumentSplitCriteria.SectionBreak;

        // Configure HtmlSaveOptions to use the split criteria.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = splitCriteria
        };

        // Save the document; Aspose.Words will create separate HTML files for each section.
        string mainHtmlPath = Path.Combine(outputDir, "SplitDocument.html");
        doc.Save(mainHtmlPath, saveOptions);

        // Validate that the split parts were created.
        string partHtmlPath = Path.Combine(outputDir, "SplitDocument-01.html");
        if (!File.Exists(mainHtmlPath) || !File.Exists(partHtmlPath))
        {
            throw new Exception("Expected split HTML files were not created.");
        }

        // Indicate successful execution.
        Console.WriteLine("Document split completed successfully.");
    }
}
