using System;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace DocumentComparisonExample
{
    class Program
    {
        static void Main()
        {
            // Create the original document in memory.
            Document originalDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(originalDoc);
            builder.Writeln("This is the original document.");
            builder.Writeln("It contains a few lines of text.");
            builder.Writeln("The quick brown fox jumps over the lazy dog.");

            // Create the edited document in memory.
            Document editedDoc = new Document();
            DocumentBuilder editedBuilder = new DocumentBuilder(editedDoc);
            editedBuilder.Writeln("This is the original document.");
            editedBuilder.Writeln("It contains a few lines of text that have been modified.");
            editedBuilder.Writeln("The quick brown fox jumps over the lazy cat.");

            // Perform the comparison. All differences will be recorded as revisions in originalDoc.
            originalDoc.Compare(editedDoc, "JD", DateTime.Now);

            // Save the comparison result. The saved file will contain the revisions (tracked changes).
            originalDoc.Save("ComparisonResult.docx");

            Console.WriteLine("Comparison completed. Result saved to ComparisonResult.docx");
        }
    }
}
