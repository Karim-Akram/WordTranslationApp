using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;

public class HomeController : Controller
{
    private readonly IHttpClientFactory _httpClientFactory;
    private const int MaxTextLength = 500; // Max length for the API query
    private const int MaxRetries = 5; // Number of retries for rate limit errors
    private const int RetryDelayMs = 1000; // Initial delay in milliseconds

    public HomeController(IHttpClientFactory httpClientFactory)
    {
        _httpClientFactory = httpClientFactory;
    }

    public IActionResult Index()
    {
        return View(); 
    }

    [HttpPost]
    public async Task<IActionResult> Upload(IFormFile file)
    {
        if (file == null || file.Length == 0)
            return Content("File not selected");

        // Translate text and create the new document
        var (translatedText, structure) = await TranslateWordFile(file);

        // Save the new document with translated text
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "translated.docx");
        CreateWordDocument(structure, translatedText, outputPath);

        return PhysicalFile(outputPath, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "translated.docx");
    }

    private async Task<(string, List<ParagraphStructure>)> TranslateWordFile(IFormFile file)
    {
        using (var stream = file.OpenReadStream())
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(stream, false))
            {
                var body = wordDoc.MainDocumentPart.Document.Body;
                var structure = new List<ParagraphStructure>();

                // Extract text and formatting
                foreach (var paragraph in body.Elements<Paragraph>())
                {
                    var paragraphStructure = new ParagraphStructure
                    {
                        Runs = paragraph.Elements<Run>()
                            .Select(run => new RunStructure
                            {
                                Text = run.InnerText,
                                RunProperties = run.RunProperties?.CloneNode(true) as RunProperties
                            })
                            .ToList()
                    };

                    structure.Add(paragraphStructure);
                }

                var combinedText = string.Join(" ", structure.SelectMany(p => p.Runs.Select(r => r.Text)));
                var translatedText = await TranslateText(combinedText);

                return (translatedText, structure);
            }
        }
    }

    private async Task<string> TranslateText(string text)
    {
        var client = _httpClientFactory.CreateClient();
        var translatedChunks = new List<string>();

        for (int i = 0; i < text.Length; i += MaxTextLength)
        {
            var chunk = text.Substring(i, Math.Min(MaxTextLength, text.Length - i));
            var response = await PostTranslation(chunk, client);

            var responseContent = await response.Content.ReadAsStringAsync();
            var translationResponse = JsonSerializer.Deserialize<TranslationResponse>(responseContent);
            translatedChunks.Add(translationResponse?.responseData?.translatedText ?? string.Empty);
        }

        return string.Join(" ", translatedChunks);
    }

    private async Task<HttpResponseMessage> PostTranslation(string text, HttpClient client)
    {
        var request = CreateTranslationRequest(text);
        return await SendRequestWithRetry(request, client);
    }

    private async Task<HttpResponseMessage> SendRequestWithRetry(HttpRequestMessage request, HttpClient client)
    {
        int retries = 0;

        while (true)
        {
            HttpResponseMessage response = await client.SendAsync(request);

            if (response.IsSuccessStatusCode)
            {
                return response;
            }

            if (response.StatusCode == HttpStatusCode.TooManyRequests)
            {
                if (retries >= MaxRetries)
                {
                    throw new HttpRequestException("Too many requests. Please try again later.");
                }

                retries++;
                await Task.Delay(RetryDelayMs * (int)Math.Pow(2, retries)); // Exponential backoff

                // Create a new request for the retry attempt
                request = CreateTranslationRequest(await request.Content.ReadAsStringAsync());
            }
            else
            {
                response.EnsureSuccessStatusCode(); // For other status codes, ensure success
                return response;
            }
        }
    }

    private HttpRequestMessage CreateTranslationRequest(string text)
    {
        return new HttpRequestMessage(HttpMethod.Post, "https://api.mymemory.translated.net/get")
        {
            Content = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string, string>("q", text),
                new KeyValuePair<string, string>("langpair", "en|ar")
            })
        };
    }

    private void CreateWordDocument(List<ParagraphStructure> structure, string translatedText, string filePath)
    {
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(filePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            Body body = mainPart.Document.AppendChild(new Body());

            var textIndex = 0;

            foreach (var paragraphStructure in structure)
            {
                var newParagraph = body.AppendChild(new Paragraph());

                foreach (var runStructure in paragraphStructure.Runs)
                {
                    var newRun = newParagraph.AppendChild(new Run());

                    // Extract and apply formatting
                    if (runStructure.RunProperties != null)
                    {
                        newRun.RunProperties = runStructure.RunProperties.CloneNode(true) as RunProperties;
                    }

                    // Apply translated text
                    var text = new Text(translatedText.Substring(textIndex, Math.Min(MaxTextLength, translatedText.Length - textIndex)));
                    newRun.AppendChild(text);

                    textIndex += text.Text.Length;
                }
            }
        }
    }

    public class TranslationResponse
    {
        public ResponseData responseData { get; set; }
    }

    public class ResponseData
    {
        public string translatedText { get; set; }
    }

    // Helper classes to store structure information
    public class ParagraphStructure
    {
        public List<RunStructure> Runs { get; set; }
    }

    public class RunStructure
    {
        public string Text { get; set; }
        public RunProperties RunProperties { get; set; }
    }
}
