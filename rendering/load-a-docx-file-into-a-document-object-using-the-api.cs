using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Path of the sample DOCX that will be created.
        string sourcePath = Path.Combine(artifactsDir, "Sample.docx");

        // -----------------------------------------------------------------
        // Create a simple DOCX file locally.
        // -----------------------------------------------------------------
        Document docToCreate = new Document();
        DocumentBuilder builder = new DocumentBuilder(docToCreate);
        builder.Writeln("Hello Aspose.Words!");
        docToCreate.Save(sourcePath);

        // -----------------------------------------------------------------
        // Load the DOCX file into a Document object using the API.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // Simple validation that the document was loaded correctly.
        // Use Contains to avoid issues with hidden paragraph marks.
        string loadedText = loadedDoc.GetText();
        if (!loadedText.Contains("Hello Aspose.Words!"))
            throw new InvalidOperationException("The loaded document does not contain the expected text.");

        // (Optional) Save the loaded document to a new file to prove it can be saved again.
        string copyPath = Path.Combine(artifactsDir, "LoadedCopy.docx");
        loadedDoc.Save(copyPath);
    }
}
