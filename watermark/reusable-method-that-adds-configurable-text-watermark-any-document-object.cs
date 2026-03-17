using System;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace WatermarkUtility
{
    /// <summary>
    /// Provides helper methods for adding text watermarks to Aspose.Words documents.
    /// </summary>
    public static class WatermarkHelper
    {
        /// <summary>
        /// Adds a text watermark to the specified <see cref="Document"/>.
        /// </summary>
        /// <param name="document">The document to which the watermark will be applied.</param>
        /// <param name="text">The watermark text. Must be non‑null, non‑empty and between 1 and 200 characters.</param>
        /// <param name="options">
        /// Optional <see cref="TextWatermarkOptions"/> that defines appearance of the watermark.
        /// If <c>null</c>, default options are used (font family = Calibri, size = auto, color = Silver, layout = Diagonal, semi‑transparent = true).
        /// </param>
        public static void AddTextWatermark(Document document, string text, TextWatermarkOptions options = null)
        {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (text == null) throw new ArgumentNullException(nameof(text));

            // If the caller did not provide options, create a default instance.
            // The default constructor already sets the library defaults, but we can
            // explicitly set common values here for clarity.
            if (options == null)
            {
                options = new TextWatermarkOptions
                {
                    FontFamily = "Calibri",          // Default font family.
                    FontSize = 0,                    // Auto size.
                    Color = System.Drawing.Color.Silver,
                    Layout = WatermarkLayout.Diagonal,
                    IsSemitrasparent = true
                };
            }

            // Use the Watermark API provided by Aspose.Words to set the text.
            // This follows the documented SetText(string, TextWatermarkOptions) method.
            document.Watermark.SetText(text, options);
        }

        /// <summary>
        /// Removes any existing watermark from the document.
        /// </summary>
        /// <param name="document">The document from which to remove the watermark.</param>
        public static void RemoveWatermark(Document document)
        {
            if (document == null) throw new ArgumentNullException(nameof(document));

            // Check the current watermark type before attempting removal.
            if (document.Watermark.Type == WatermarkType.Text ||
                document.Watermark.Type == WatermarkType.Image)
            {
                document.Watermark.Remove();
            }
        }
    }

    // Example usage.
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Configure custom watermark options (optional).
            TextWatermarkOptions customOptions = new TextWatermarkOptions
            {
                FontFamily = "Arial",
                FontSize = 36,
                Color = System.Drawing.Color.Black,
                Layout = WatermarkLayout.Diagonal,
                IsSemitrasparent = false
            };

            // Add the watermark.
            WatermarkHelper.AddTextWatermark(doc, "Confidential", customOptions);

            // Save the document (using the standard Save method as per lifecycle rules).
            doc.Save("OutputWithWatermark.docx");
        }
    }
}
