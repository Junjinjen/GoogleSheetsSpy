using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Media;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace GoogleSheetsSpy
{
    internal static class Program
    {
        private const string SettingsFilename = "appsettings.json";
        private const string CookiesHeader = "cookie";
        private const string CookiesConfigurationKey = "Cookies";
        private const string SpreadsheetIdConfigurationKey = "SpreadsheetId";
        private const string RequestsDelayConfigurationKey = "RequestsDelay";
        private const string SoundFilenameConfigurationKey = "SoundFilename";
        private const string GoogleSheetDownloadLink = @"https://docs.google.com/spreadsheets/d/{0}/export?format=xlsx";

        private const string WorksheetName = "1361 урок";
        private const string LessonNumberColumn = "F";
        private const string AssignedWorkerColumn = "L";

        private const string GreenColor = "FF00FF00";
        private static readonly List<ICellTrigger> FinalTriggers = new()
        {
            new StyledTextTrigger { Text = "финал", BackgroundColor = GreenColor, },
            new StyledTextTrigger { Text = "фінал", BackgroundColor = GreenColor, },
        };

        private static readonly List<ICellTrigger> ReadyTriggers = new()
        {
            new StyledTextTrigger { Text = "Готово", BackgroundColor = GreenColor, },
        };

        private static readonly List<ICellTrigger> ToTheOnlineSchoolTriggers = new()
        {
            new StyledTextTrigger { Text = "на вшо", BackgroundColor = GreenColor, },
            new StyledTextTrigger { Text = "финал", BackgroundColor = GreenColor, },
            new StyledTextTrigger { Text = "фінал", BackgroundColor = GreenColor, },
        };

        private static readonly List<ICellTrigger> From7to11Triggers = new()
        {
            new TextTrigger { Text = "7", },
            new TextTrigger { Text = "8", },
            new TextTrigger { Text = "9", },
            new TextTrigger { Text = "10", },
            new TextTrigger { Text = "11", },
        };

        private static readonly List<ICellTrigger> From5to6Triggers = new()
        {
            new TextTrigger { Text = "5", },
            new TextTrigger { Text = "6", },
            new TextTrigger { Text = "5-6", },
        };

        private static readonly List<ICellTrigger> LanguageMathTriggers = new()
        {
            new TextTrigger { Text = "Математика", },
            new TextTrigger { Text = "Українська мова", },
        };

        private static readonly List<ICellTrigger> HistoryTriggers = new()
        {
            new TextTrigger { Text = "Історія України", },
            new TextTrigger { Text = "Історія та Громадянська освіта", },
            new TextTrigger { Text = "Історія України та Громадянська Освіта", },
        };

        private static readonly List<Dictionary<string, List<ICellTrigger>>> TriggerGroups = new()
        {
            // 7-11 класс
            new()
            {
                { "A", From7to11Triggers }, // 7-11 класс
                { "H", FinalTriggers }, // Конспект
                { "I", FinalTriggers }, // Тест
                //{ "J", FinalTriggers }, // Джерела
                { "Q", ToTheOnlineSchoolTriggers }, // Видео
                //{ "U", ReadyTriggers }, // Субтитры
            },
            new()
            {
                { "A", From7to11Triggers }, // 7-11 класс
                //{ "H", FinalTriggers }, // Конспект
                { "I", FinalTriggers }, // Тест
                { "J", FinalTriggers }, // Джерела
                { "Q", ToTheOnlineSchoolTriggers }, // Видео
                //{ "U", ReadyTriggers }, // Субтитры
            },

            // 5-6 класс, математика и укр. яз.
            new()
            {
                { "A", From5to6Triggers }, // 5-6 класс
                { "B", LanguageMathTriggers }, // Урок
                { "H", FinalTriggers }, // Конспект
                { "I", FinalTriggers }, // Тест
                //{ "J", FinalTriggers }, // Джерела
                { "Q", ToTheOnlineSchoolTriggers }, // Видео
                //{ "U", ReadyTriggers }, // Субтитры
            },
            new()
            {
                { "A", From5to6Triggers }, // 5-6 класс
                { "B", LanguageMathTriggers }, // Урок
                //{ "H", FinalTriggers }, // Конспект
                { "I", FinalTriggers }, // Тест
                { "J", FinalTriggers }, // Джерела
                { "Q", ToTheOnlineSchoolTriggers }, // Видео
                //{ "U", ReadyTriggers }, // Субтитры
            },

            // 5-6 класс, история
            new()
            {
                { "A", From5to6Triggers }, // 5-6 класс
                { "B", HistoryTriggers }, // Урок
                { "H", FinalTriggers }, // Конспект
                { "I", FinalTriggers }, // Тест
                { "J", FinalTriggers }, // Джерела
                { "Q", ToTheOnlineSchoolTriggers }, // Видео
                //{ "U", ReadyTriggers }, // Субтитры
            },
        };

        private static readonly IConfigurationRoot Configuration = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile(SettingsFilename, optional: false, reloadOnChange: true)
            .Build();

        private static readonly SoundPlayer SoundPlayer = new(Configuration[SoundFilenameConfigurationKey]);

        static Program()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public async static Task Main(string[] args)
        {
            var requestsDelay = Configuration.GetValue<int>(RequestsDelayConfigurationKey);

            while (true)
            {
                try
                {
                    using var stream = await GetSheetStream();
                    CheckSheet(stream);

                    Thread.Sleep(requestsDelay);
                }
                catch (HttpRequestException e)
                {
                    if (e.StatusCode == HttpStatusCode.TooManyRequests)
                    {
                        Thread.Sleep(requestsDelay);
                        continue;
                    }

                    if (e.StatusCode == HttpStatusCode.BadGateway)
                    {
                        Thread.Sleep(requestsDelay);
                        continue;
                    }

                    throw;
                }
            }
        }

        private static void CheckSheet(Stream sheetStream)
        {
            using var package = new ExcelPackage(sheetStream);
            var worksheet = package.Workbook.Worksheets.First(x => x.Name == WorksheetName);
            var shouldPlay = false;

            Console.Clear();
            for (int row = 1; row <= worksheet.Dimension.Rows; row++)
            {
                var cells = worksheet.Cells[$"{row}:{row}"];
                if (!string.IsNullOrEmpty(cells[$"{AssignedWorkerColumn}{row}"].Text))
                {
                    continue;
                }

                var result = TriggerGroups.Any(x => x.All(y => y.Value.Any(z => z.IsTriggered(cells[$"{y.Key}{row}"]))));
                if (!result)
                {
                    continue;
                }

                shouldPlay = true;
                var lessonNumber = cells[$"{LessonNumberColumn}{row}"].Text;
                Console.WriteLine($"Бягом забирай №{lessonNumber}");
            }

            if (shouldPlay)
            {
                SoundPlayer.Play();
            }
        }

        private async static Task<Stream> GetSheetStream()
        {
            using var client = new HttpClient();
            client.DefaultRequestHeaders.Add(CookiesHeader, Configuration[CookiesConfigurationKey]);

            var requestUrl = string.Format(GoogleSheetDownloadLink, Configuration[SpreadsheetIdConfigurationKey]);
            var response = await client.GetAsync(requestUrl);
            response.EnsureSuccessStatusCode();

            return await response.Content.ReadAsStreamAsync();
        }
    }
}