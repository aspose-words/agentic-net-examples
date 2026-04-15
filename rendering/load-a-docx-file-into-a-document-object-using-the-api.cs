using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define folders for the sample source and the output.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");

        // Ensure the folders exist.
        Directory.CreateDirectory(dataDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a simple DOCX file that will be used as the source.
        // -----------------------------------------------------------------
        string sourcePath = Path.Combine(dataDir, "Sample.docx");
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Hello Aspose.Words!");
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the DOCX file into a Document object using the API.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // Verify that the document was loaded correctly.
        // Use Contains to avoid issues with line‑break characters.
        if (!loadedDoc.GetText().Contains("Hello Aspose.Words!"))
            throw new InvalidOperationException("The document was not loaded as expected.");

        // -----------------------------------------------------------------
        // 3. Save the loaded document to demonstrate that it can be used.
        // -----------------------------------------------------------------
        string resultPath = Path.Combine(outputDir, "Loaded.docx");
        loadedDoc.Save(resultPath);
    }
}
