using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create the original document (portrait orientation).
        Document docOriginal = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(docOriginal);
        builderOriginal.PageSetup.Orientation = Orientation.Portrait;
        builderOriginal.Writeln("This paragraph is used to test page orientation changes.");

        // Create the edited document (landscape orientation) with the same content.
        Document docEdited = new Document();
        DocumentBuilder builderEdited = new DocumentBuilder(docEdited);
        builderEdited.PageSetup.Orientation = Orientation.Landscape;
        builderEdited.Writeln("This paragraph is used to test page orientation changes.");

        // Ensure both documents have no revisions before comparison.
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
        {
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");
        }

        // Compare the documents. The original document will receive revisions.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now);

        // Verify that revisions were created.
        int totalRevisions = docOriginal.Revisions.Count;
        int formatRevisions = docOriginal.Revisions.Count(r => r.RevisionType == RevisionType.FormatChange);

        Console.WriteLine($"Total revisions detected: {totalRevisions}");
        Console.WriteLine($"Format (orientation) revisions detected: {formatRevisions}");

        // Save the comparison result with revisions.
        string resultPath = Path.Combine(artifactsDir, "OrientationComparisonResult.docx");
        docOriginal.Save(resultPath);
    }
}
