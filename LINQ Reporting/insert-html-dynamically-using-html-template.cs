using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // HTML template with placeholders that will be replaced at runtime.
        string htmlTemplate = @"
            <html>
                <body>
                    <h1>Hello, {Name}!</h1>
                    <p>Welcome to <b>{Company}</b>.</p>
                </body>
            </html>";

        // Values to insert into the template.
        string name = "John Doe";
        string company = "Acme Corp";

        // Replace placeholders with actual data.
        string html = htmlTemplate
            .Replace("{Name}", name)
            .Replace("{Company}", company);

        // Insert the processed HTML into the document.
        // Use builder formatting as the base formatting for the inserted HTML.
        builder.InsertHtml(html, HtmlInsertOptions.UseBuilderFormatting);

        // Save the resulting document.
        doc.Save("DynamicHtml.docx");
    }
}
