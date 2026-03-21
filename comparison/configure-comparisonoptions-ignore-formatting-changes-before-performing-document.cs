using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

class ContractComparison
{
    static void Main()
    {
        // Create a temporary folder for the demo files.
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeComparisonDemo");
        Directory.CreateDirectory(tempFolder);

        // Paths to the original, revised, and result documents.
        string originalPath = Path.Combine(tempFolder, "OriginalContract.docx");
        string revisedPath  = Path.Combine(tempFolder, "RevisedContract.docx");
        string resultPath   = Path.Combine(tempFolder, "ContractComparisonResult.docx");

        // Build the original document.
        Document originalDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(originalDoc);
        builder.Writeln("This is the original contract.");
        builder.Font.Bold = true;
        builder.Writeln("The price is $1000.");
        originalDoc.Save(originalPath);

        // Build the revised document (with a content change and a formatting change).
        Document revisedDoc = new Document();
        builder = new DocumentBuilder(revisedDoc);
        builder.Writeln("This is the original contract.");
        builder.Font.Bold = false; // formatting change that we will ignore
        builder.Writeln("The price is $1200."); // content change
        revisedDoc.Save(revisedPath);

        // Load the two documents that need to be compared.
        Document docOriginal = new Document(originalPath);
        Document docRevised  = new Document(revisedPath);

        // Configure comparison options to ignore all formatting changes.
        CompareOptions compareOptions = new CompareOptions
        {
            IgnoreFormatting = true,
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. All formatting differences will be ignored.
        docOriginal.Compare(docRevised, "LegalTeam", DateTime.Now, compareOptions);

        // Save the document that now contains revision marks for the content differences.
        docOriginal.Save(resultPath);

        Console.WriteLine($"Comparison completed. Result saved to: {resultPath}");
    }
}
