using System;
using System.IO;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Comparing;

public class DetectMovedParagraphsExample
{
    public static void Main()
    {
        // Create the original document with two paragraphs.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("First paragraph.");
        builderOriginal.Writeln("Second paragraph that will be moved.");

        // Save the original document to a temporary file (optional, just for illustration).
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "original.docx");
        original.Save(originalPath);

        // Create the revised document where the second paragraph is moved before the first.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Second paragraph that will be moved."); // Moved paragraph.
        builderRevised.Writeln("First paragraph.");

        // Save the revised document to a temporary file (optional).
        string revisedPath = Path.Combine(Directory.GetCurrentDirectory(), "revised.docx");
        revised.Save(revisedPath);

        // Set up comparison options.
        // The DetectMove flag may not exist in older versions of Aspose.Words.
        // Use reflection to set it when available, avoiding compile‑time errors.
        CompareOptions compareOptions = new CompareOptions();
        PropertyInfo? detectMoveProp = typeof(CompareOptions).GetProperty("DetectMove", BindingFlags.Public | BindingFlags.Instance);
        if (detectMoveProp != null && detectMoveProp.CanWrite)
        {
            detectMoveProp.SetValue(compareOptions, true);
        }

        // Perform the comparison. Revisions will be added to the original document.
        original.Compare(revised, "Comparer", DateTime.Now, compareOptions);

        // Count move‑related revisions.
        int moveFromCount = 0;
        int moveToCount = 0;

        foreach (Revision revision in original.Revisions)
        {
            // Use the string representation of the enum to avoid referencing members that may not exist.
            string revTypeName = revision.RevisionType.ToString();
            if (revTypeName.Equals("MoveFrom", StringComparison.OrdinalIgnoreCase))
                moveFromCount++;
            else if (revTypeName.Equals("MoveTo", StringComparison.OrdinalIgnoreCase))
                moveToCount++;
        }

        // Output the results.
        Console.WriteLine($"MoveFrom revisions detected: {moveFromCount}");
        Console.WriteLine($"MoveTo   revisions detected: {moveToCount}");
        Console.WriteLine($"Total revisions detected: {original.Revisions.Count}");

        // Save the compared document showing the revisions.
        string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "comparison-result.docx");
        original.Save(resultPath);
    }
}
