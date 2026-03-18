using System;
using Aspose.Words;

namespace WatermarkValidationExample
{
    public class WatermarkValidator
    {
        // Validates that the document at the specified path does not contain a watermark.
        // Throws an InvalidOperationException if a watermark is present.
        public static void ValidateNoWatermark(string documentPath)
        {
            // Load the document using Aspose.Words Document constructor.
            Document doc = new Document(documentPath);

            // Check the watermark type. WatermarkType.None indicates that no watermark is set.
            if (doc.Watermark.Type != WatermarkType.None)
            {
                // A watermark exists; raise an exception to prevent publishing.
                throw new InvalidOperationException("The document contains a watermark and cannot be published.");
            }
        }
    }

    class Program
    {
        // Entry point required by the C# compiler for a console application.
        static void Main(string[] args)
        {
            // Simple argument handling – first argument is the document path.
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: WatermarkValidationExample <documentPath>");
                return;
            }

            string documentPath = args[0];
            try
            {
                WatermarkValidator.ValidateNoWatermark(documentPath);
                Console.WriteLine("Document validation succeeded – no watermark found.");
            }
            catch (InvalidOperationException ex)
            {
                Console.WriteLine($"Validation failed: {ex.Message}");
                // In a real publishing pipeline you would abort further processing here.
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
