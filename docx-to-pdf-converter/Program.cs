using Microsoft.Extensions.Configuration;
using Microsoft.Identity.Client;
using System.Net.Http.Headers;

try
{
    string blobUrl = "https://vouchstoragedev.blob.core.windows.net/roa-templates/zest-life-contact-roa-template.docx";
    var outputPdfPath = "output.pdf";

    using var httpClient = new HttpClient();
    var docBytes = await httpClient.GetByteArrayAsync(blobUrl);

    // Build configuration
    var config = new ConfigurationBuilder()
        .AddJsonFile("appsettings.json")
        .Build();

    var tenantId = config["AzureAd:TenantId"];
    var clientId = config["AzureAd:ClientId"];
    var clientSecret = config["AzureAd:ClientSecret"];
    var userPrincipal = config["AzureAd:UserPrincipal"];

    var cca = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithClientSecret(clientSecret)
                .WithAuthority($"https://login.microsoftonline.com/{tenantId}/")
                .Build();

    var tokenResult = await cca
                .AcquireTokenForClient(["https://graph.microsoft.com/.default"])
                .ExecuteAsync();

    using var http = new HttpClient();

    var request = new HttpRequestMessage(
        HttpMethod.Put,
        $"https://graph.microsoft.com/v1.0/users/{userPrincipal}/drive/root:/temp.docx:/content?format=pdf"
    );

    request.Headers.Authorization =
        new AuthenticationHeaderValue("Bearer", tokenResult.AccessToken);

    request.Content = new ByteArrayContent(docBytes);

    request.Content.Headers.ContentType =
        new MediaTypeHeaderValue(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        );

    var response = await http.SendAsync(request);

    response.EnsureSuccessStatusCode();

    var pdfBytes = await response.Content.ReadAsByteArrayAsync();

    await File.WriteAllBytesAsync(outputPdfPath, pdfBytes);
}
catch (Exception ex)
{
    throw new Exception("An error occurred while processing the document.", ex);
}