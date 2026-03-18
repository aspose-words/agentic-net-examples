using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeWordsLinkExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Example URLs.
            string url1 = "https://www.aspose.com/";
            string url2 = "https://docs.aspose.com/words/net/";

            // Insert a hyperlink with explicit display text.
            InsertLink(builder, url1, "Aspose Home");

            // Insert a hyperlink without providing display text – it will default to the URL.
            InsertLink(builder, url2);

            // Save the document.
            doc.Save("Hyperlinks.docx");
        }

        /// <summary>
        /// Inserts a hyperlink into the document. If <paramref name="displayText"/> is null or empty,
        /// the method uses the <paramref name="url"/> as the visible text.
        /// </summary>
        /// <param name="builder">The DocumentBuilder used to insert the field.</param>
        /// <param name="url">The target URL of the hyperlink.</param>
        /// <param name="displayText">Optional text to display for the link.</param>
        private static void InsertLink(DocumentBuilder builder, string url, string displayText = null)
        {
            // Default the display text to the URL when it is not supplied.
            if (string.IsNullOrEmpty(displayText))
                displayText = url;

            // Set hyperlink formatting (blue and underlined) – required because InsertHyperlink does not apply formatting automatically.
            builder.Font.Color = Color.Blue;
            builder.Font.Underline = Underline.Single;

            // Insert the hyperlink field. The third argument 'false' indicates that the second argument is a URL, not a bookmark.
            Field field = builder.InsertHyperlink(displayText, url, false);

            // Optionally, you can cast to FieldHyperlink to set additional properties such as ScreenTip.
            if (field is FieldHyperlink hyperlink)
                hyperlink.ScreenTip = $"Open {url}";

            // Reset formatting after the hyperlink so subsequent text is not affected.
            builder.Font.ClearFormatting();
        }
    }
}
