using System;
using Aspose.Words;
using Aspose.Words.AI;

class Program
{
    static void Main()
    {
        // Create a blank Word document.
        Document doc = new Document();

        // Use DocumentBuilder to add some sample text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world! This is a sample document created with Aspose.Words.");

        // Save the original document.
        doc.Save("Original.docx");

        // Retrieve the API key for the AI service from environment variables.
        string apiKey = Environment.GetEnvironmentVariable("API_KEY");

        // Create an AI model instance (Gemini Flash) and set the API key.
        AiModel model = AiModel.Create(AiModelType.GeminiFlashLatest).WithApiKey(apiKey);

        // Translate the document to French using the AI model.
        Document translatedDoc = model.Translate(doc, Language.French);

        // Save the translated document.
        translatedDoc.Save("Translated_French.docx");
    }
}
