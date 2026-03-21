using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields;

class PreserveOfficeMath
{
    static void Main()
    {
        const string inputPath = "InputWithMath.docx";
        const string outputPath = "OutputPreservedMath.docx";

        Document doc;

        if (File.Exists(inputPath))
        {
            // Load the existing document that contains OfficeMath equations.
            doc = new Document(inputPath);
        }
        else
        {
            // Create a placeholder document when the input file is missing.
            doc = new Document();
            DocumentBuilder placeholderBuilder = new DocumentBuilder(doc);
            placeholderBuilder.Writeln("Placeholder document created because InputWithMath.docx was not found.");
            // Optionally, add a simple OfficeMath equation.
            placeholderBuilder.InsertField("EQ \\o\\ac(\\up\\a\\to\\b)");
        }

        // Example modification: add a paragraph at the end of the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Additional content added while preserving equations.");

        // Prepare save options for DOCX format.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            SaveFormat = SaveFormat.Docx // Ensure the output is DOCX.
        };

        // Save the modified document.
        doc.Save(outputPath, saveOptions);
    }
}
