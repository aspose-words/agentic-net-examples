using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Input DOT template file path
        string inputPath = "Template.dot";

        // Desired output MHTML file path
        string outputPath = "Result.mht";

        // Load the DOT document
        Document doc = new Document(inputPath);

        // Save the document in MHTML format
        doc.Save(outputPath, SaveFormat.Mhtml);
    }
}
