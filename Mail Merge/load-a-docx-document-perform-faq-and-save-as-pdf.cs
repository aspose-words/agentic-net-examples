using System;
using Aspose.Words;
using Aspose.Words.AI;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = "input.docx";

        // Path where the resulting PDF will be saved.
        string outputPath = "output.pdf";

        // Load the DOCX document using the Document constructor (load rule).
        Document doc = new Document(inputPath);

        // Create an AI model instance (e.g., OpenAI GPT‑4o mini). 
        // The model will be used to generate a summary/FAQ of the document.
        AiModel model = AiModel.Create(AiModelType.Gpt4OMini);

        // Prepare summarization options (optional – can be left with defaults).
        SummarizeOptions options = new SummarizeOptions();

        // Generate a summarized version of the document (FAQ) using the Summarize method (AI rule).
        Document faqDoc = model.Summarize(doc, options);

        // Save the summarized document as PDF using the Save method (save rule).
        faqDoc.Save(outputPath);
    }
}
