using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Folder that contains the source DOCX files.
        string sourceFolder = @"Docs";
        // Folder where HTML files and extracted images will be placed.
        string outputFolder = @"Output";

        Directory.CreateDirectory(sourceFolder);
        Directory.CreateDirectory(outputFolder);

        // Collect information for the index page.
        var indexEntries = new List<IndexEntry>();

        foreach (string docPath in Directory.GetFiles(sourceFolder, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(docPath);

            string docName = Path.GetFileNameWithoutExtension(docPath);
            string htmlFilePath = Path.Combine(outputFolder, docName + ".html");
            string imagesFolderPath = Path.Combine(outputFolder, docName + "_files");
            string imagesFolderAlias = docName + "_files";

            // Ensure the images folder exists.
            Directory.CreateDirectory(imagesFolderPath);

            // Save as HTML and let Aspose.Words extract images to the specified folder.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ImagesFolder = imagesFolderPath,
                ImagesFolderAlias = imagesFolderAlias
            };
            doc.Save(htmlFilePath, htmlOptions);

            // Determine the first image for thumbnail (if any).
            string thumbnailPath = null;
            string[] imageFiles = Directory.GetFiles(imagesFolderPath);
            if (imageFiles.Length > 0)
                thumbnailPath = Path.GetFileName(imageFiles[0]); // relative to HTML file

            indexEntries.Add(new IndexEntry
            {
                Title = docName,
                HtmlFile = Path.GetFileName(htmlFilePath),
                ThumbnailFile = thumbnailPath,
                ImagesFolderAlias = imagesFolderAlias
            });
        }

        // Build the index HTML.
        var sb = new StringBuilder();
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html lang=\"en\">");
        sb.AppendLine("<head><meta charset=\"UTF-8\"><title>Document Index</title></head>");
        sb.AppendLine("<body>");
        sb.AppendLine("<h1>Document Index</h1>");
        sb.AppendLine("<ul>");

        foreach (var entry in indexEntries)
        {
            sb.AppendLine("<li>");
            sb.AppendLine($"<a href=\"{entry.HtmlFile}\">{entry.Title}</a>");
            if (!string.IsNullOrEmpty(entry.ThumbnailFile))
            {
                string imgSrc = $"{entry.ImagesFolderAlias}/{entry.ThumbnailFile}";
                sb.AppendLine($"<br><img src=\"{imgSrc}\" alt=\"{entry.Title}\" style=\"max-width:200px;\"/>");
            }
            sb.AppendLine("</li>");
        }

        sb.AppendLine("</ul>");
        sb.AppendLine("</body>");
        sb.AppendLine("</html>");

        // Save the index page.
        string indexPath = Path.Combine(outputFolder, "index.html");
        File.WriteAllText(indexPath, sb.ToString(), Encoding.UTF8);
    }

    // Simple DTO for index generation.
    class IndexEntry
    {
        public string Title { get; set; }
        public string HtmlFile { get; set; }
        public string ThumbnailFile { get; set; }
        public string ImagesFolderAlias { get; set; }
    }
}
