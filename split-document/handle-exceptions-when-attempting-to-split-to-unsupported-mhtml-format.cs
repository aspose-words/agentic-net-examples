using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with two pages.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("This is the first page.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is the second page.");

        // Save the source document locally.
        string sourcePath = "Sample.docx";
        doc.Save(sourcePath);

        // Attempt to split the first page and save it as MHTML.
        try
        {
            var sourceDoc = new Document(sourcePath);
            // Extract the first page (page numbers are 1‑based).
            var splitDoc = sourceDoc.ExtractPages(1, 1);

            string mhtmlPath = "Page1.mhtml";
            // Saving to MHTML format. This may throw if the format does not support page‑based splitting.
            splitDoc.Save(mhtmlPath, SaveFormat.Mhtml);

            // Verify that the file was created.
            if (File.Exists(mhtmlPath))
            {
                Console.WriteLine($"Split page saved successfully as MHTML: {mhtmlPath}");
            }
            else
            {
                throw new FileNotFoundException("Expected MHTML file was not created.", mhtmlPath);
            }
        }
        catch (Exception ex)
        {
            // Handle any exception that occurs during the split or save operation.
            Console.WriteLine($"Exception occurred while attempting to split to MHTML: {ex.Message}");
        }
    }
}
