using System;
using Aspose.Words;
using Aspose.Words.AI;

class Program
{
    static void Main()
    {
        // Create a blank Word document.
        Document originalDoc = new Document();

        // Use DocumentBuilder to add natural language specification text.
        DocumentBuilder builder = new DocumentBuilder(originalDoc);
        builder.Writeln("This document contains a natural language specification.");
        builder.Writeln("It describes the requirements, constraints, and expected outcomes for the project.");
        builder.Writeln("All sections are written in clear, concise English.");

        // Save the original document.
        originalDoc.Save("OriginalSpecification.docx");

        // Translate the document to French using a Google AI model.
        // The GoogleAiModel constructor requires an API key (or configuration) – replace the placeholder with a valid key.
        var aiModel = new GoogleAiModel("YOUR_GOOGLE_API_KEY");

        // Translate the document. The Language enum defines target languages.
        Document translatedDoc = aiModel.Translate(originalDoc, Language.French);

        // Save the translated document.
        translatedDoc.Save("TranslatedSpecification_French.docx");
    }
}
