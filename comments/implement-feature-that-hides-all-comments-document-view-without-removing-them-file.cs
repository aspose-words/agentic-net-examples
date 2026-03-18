using System;
using Aspose.Words;
using Aspose.Words.Layout;

public class CommentHider
{
    /// <summary>
    /// Loads a Word document, hides all comments in the view, and saves the result.
    /// The comments remain in the file; they are only not rendered.
    /// </summary>
    /// <param name="inputFile">Path to the source document.</param>
    /// <param name="outputFile">Path where the modified document will be saved.</param>
    public static void HideAllComments(string inputFile, string outputFile)
    {
        // Load the existing document.
        Document doc = new Document(inputFile);

        // Set the comment rendering mode to Hide.
        // This prevents comments from being displayed in the layout,
        // but the comment nodes stay intact in the document structure.
        doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.Hide;

        // Rebuild the layout so the change takes effect.
        doc.UpdatePageLayout();

        // Save the document (format is inferred from the file extension).
        doc.Save(outputFile);
    }

    // Example usage.
    public static void Main()
    {
        string sourcePath = @"C:\Docs\Sample.docx";
        string resultPath = @"C:\Docs\Sample_HiddenComments.docx";

        HideAllComments(sourcePath, resultPath);

        Console.WriteLine("Comments have been hidden in the saved document.");
    }
}
