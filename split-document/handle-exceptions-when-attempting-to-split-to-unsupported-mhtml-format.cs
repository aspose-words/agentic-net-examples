using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with two pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Page 1 content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 content.");

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Get the total number of pages in the document.
        int pageCount = doc.PageCount;

        // Attempt to split each page and save as MHTML.
        for (int page = 1; page <= pageCount; page++)
        {
            try
            {
                // Extract a single page as a new document.
                Document pageDoc = doc.ExtractPages(page, 1);

                // Prepare save options for MHTML using the generic factory method.
                SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Mhtml);

                // Define the output file path.
                string outPath = Path.Combine(outputDir, $"Page_{page}.mhtml");

                // Save the extracted page as MHTML.
                pageDoc.Save(outPath, saveOptions);

                // Verify that the file was created.
                if (!File.Exists(outPath))
                {
                    throw new InvalidOperationException($"Failed to create file: {outPath}");
                }
            }
            catch (NotSupportedException ex)
            {
                // Handle the case where MHTML splitting is not supported.
                Console.WriteLine($"MHTML splitting is not supported for page {page}: {ex.Message}");
            }
            catch (Exception ex)
            {
                // Handle any other unexpected exceptions.
                Console.WriteLine($"An error occurred while processing page {page}: {ex.Message}");
            }
        }
    }
}
