using System;
using System.Text;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Text.Json;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using CsvHelper;
using System.Globalization;

namespace NGP_VAN_APP
{
    class Program
    {
        static private HttpClient _httpClient = new HttpClient();
        static private string authParams = Convert.ToBase64String(Encoding.ASCII.GetBytes("part1:part2"));

        static async Task Main(string[] args)
        {
            _httpClient.BaseAddress = new Uri("https://api.myngp.com/v2/");
            _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", authParams);

            await GetReport();
        }

        private class ItemsWrapper
        {
            public IEnumerable<SentEmail> Items { get; set; }
        }

        private class SentEmail
        {
            public int EmailMessageId { get; set; }
        }        

        private class EmailDetail
        {
            public int EmailMessageId { get; set; }
            public string Name { get; set; }
            public IEnumerable<Variant> Variants { get; set; }
            public EmailStatistics Statistics { get; set; }
            public string TopVariant { get; set; }
        }

        private class Variant
        {
            public int EmailMessageVariantId { get; set; }
            public string Name { get; set; }
            public EmailStatistics Statistics { get; set; }
        }

        private class EmailStatistics
        {
            public int Recipients { get; set; }
            public int Opens { get; set; }
            public int Clicks { get; set; }
            public int Unsubscribes { get; set; }
            public int Bounces { get; set; }
        }

        static T DeserializeCaseInsensitive<T>(string json)
        {
            return JsonSerializer.Deserialize<T>(json, 
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
        }

        static async Task<string> GetEndpointJson(string endpoint)
        {
            return await _httpClient.GetStringAsync(endpoint);
        }

        static async Task GetReport()
        {
            var sentEmailsJson = await GetEndpointJson("broadcastEmails");
            var sentEmails = DeserializeCaseInsensitive<ItemsWrapper>(sentEmailsJson).Items;

            var emailDetails = (await GetEmailDetails(sentEmails)).ToList();

            for (int i = 0; i < emailDetails.Count(); i++)
            {
                emailDetails[i].TopVariant = await GetTopVariant(emailDetails[i]);
            }

            await WriteCsv(emailDetails.OrderByDescending(emailDetail => emailDetail.EmailMessageId));
        }

        static async Task<IEnumerable<EmailDetail>> GetEmailDetails(IEnumerable<SentEmail> sentEmails)
        {
            var getDetailTasks = sentEmails.Select(sentEmail => GetEndpointJson($"broadcastEmails/{sentEmail.EmailMessageId}?$expand=statistics"));
            var emailDetailsJson = await Task.WhenAll(getDetailTasks);
            return emailDetailsJson.Select(emailDetailJson => DeserializeCaseInsensitive<EmailDetail>(emailDetailJson));
        }

        static async Task<IEnumerable<Variant>> GetVariants(EmailDetail emailDetail)
        {
            var getVariantTasks = emailDetail
                .Variants
                .Select(variant => GetEndpointJson($"broadcastEmails/{emailDetail.EmailMessageId}/variants/{variant.EmailMessageVariantId}?$expand=statistics"));

            var variantsJson = await Task.WhenAll(getVariantTasks);
            return variantsJson.Select(variantJson => DeserializeCaseInsensitive<Variant>(variantJson));
        }

        static async Task<string> GetTopVariant(EmailDetail emailDetail)
        {
            var variants = await GetVariants(emailDetail);

            return variants.Count() == 1
                ? variants.First().Name
                : variants
                    .Aggregate((current, next) => (current.Statistics.Opens / current.Statistics.Recipients) > (next.Statistics.Opens / next.Statistics.Recipients) ? current : next)
                    .Name;
        }

        static async Task WriteCsv(IEnumerable<EmailDetail> emailDetail)
        {
            var fileName = $"email-report-{DateTime.Now.ToFileTime()}.csv";
            var targetPath = Path.Combine(Environment.CurrentDirectory, fileName);
            using var streamWriter = new StreamWriter(targetPath);
            using var csvWriter = new CsvWriter(streamWriter, CultureInfo.InvariantCulture);
            await csvWriter.WriteRecordsAsync(emailDetail
                .Select(emailDetail => new 
                {
                    emailDetail.EmailMessageId,
                    emailDetail.Name,
                    emailDetail.Statistics,
                    emailDetail.TopVariant
                }));

            Console.WriteLine($"{fileName} written to directory: {Environment.CurrentDirectory}");
        }
    }
}
