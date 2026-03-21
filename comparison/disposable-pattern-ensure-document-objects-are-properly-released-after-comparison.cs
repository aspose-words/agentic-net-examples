using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

namespace DocumentComparisonDemo
{
    class Program
    {
        static void Main()
        {
            // Create a temporary folder for the demo files.
            string tempFolder = Path.Combine(Path.GetTempPath(), "DocumentComparisonDemo");
            Directory.CreateDirectory(tempFolder);

            // Paths to the source documents.
            string originalPath = Path.Combine(tempFolder, "Original.docx");
            string editedPath   = Path.Combine(tempFolder, "Edited.docx");
            string resultPath   = Path.Combine(tempFolder, "ComparisonResult.docx");

            // Ensure the source documents exist.
            if (!File.Exists(originalPath))
            {
                var doc = new Document();
                var builder = new DocumentBuilder(doc);
                builder.Writeln("This is the original document.");
                doc.Save(originalPath);
            }

            if (!File.Exists(editedPath))
            {
                var doc = new Document();
                var builder = new DocumentBuilder(doc);
                builder.Writeln("This is the edited document with a small change.");
                doc.Save(editedPath);
            }

            // Use using statements to guarantee that the file streams are closed and disposed
            // as soon as the comparison operation is finished.
            using (FileStream originalStream = File.OpenRead(originalPath))
            using (FileStream editedStream   = File.OpenRead(editedPath))
            {
                // Load the documents from the streams. The Document constructors
                // accept a Stream and therefore do not keep the stream open after loading.
                Document originalDoc = new Document(originalStream);
                Document editedDoc   = new Document(editedStream);

                // Ensure both documents have no revisions before performing a comparison.
                if (originalDoc.Revisions.Count == 0 && editedDoc.Revisions.Count == 0)
                {
                    // Configure comparison options as needed.
                    CompareOptions compareOptions = new CompareOptions
                    {
                        // Example: ignore case changes and formatting differences.
                        IgnoreCaseChanges = true,
                        IgnoreFormatting  = true,
                        // Use the original document as the base for comparison.
                        Target = ComparisonTargetType.Current
                    };

                    // Perform the comparison. The revisions are added to originalDoc.
                    originalDoc.Compare(editedDoc, "Comparer", DateTime.Now, compareOptions);
                }

                // Save the comparison result. The Save method determines the format from the file extension.
                originalDoc.Save(resultPath);
                Console.WriteLine($"Comparison result saved to: {resultPath}");
            } // Both FileStream objects are disposed here; the Document instances are now eligible for GC.
        }
    }
}
