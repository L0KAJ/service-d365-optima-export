using Microsoft.Identity.Client;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;


namespace ImportDynamicsKsef
{
    public partial class Service1 : ServiceBase
    {
        System.Timers.Timer _timer = new System.Timers.Timer();
        private bool _isRunning = false;

        // Obiekt przechowywujący konfigurację
        private AppConfig _config;

        // Ścieżka do pliku konfiguracyjnego
        private readonly string _configPath = @"C:\SettingsServicesOptima\ImportDynamicsSettings.json";

        public Service1()
        {
            InitializeComponent();
        }


        protected override void OnStart(string[] args)
        {
            try
            {
                LoadConfig();

                // Tworzenie katalogów 
                Directory.CreateDirectory(_config.Paths.ExportFolder);
                Directory.CreateDirectory(Path.GetDirectoryName(_config.Paths.ErrorLog));

                // Timer
                _timer = new System.Timers.Timer();
                _timer.Elapsed += Timer_Elapsed;
                _timer.AutoReset = false; // WAŻNE: False

                // Pierwsze planowanie uruchomienia
                ScheduleNextRun();
            }
            catch (Exception ex)
            {
                // Fallback logging
                string fallbackLog = @"C:\Fallback_log_services\ImportDynamics_startup_error.txt";
                try
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(fallbackLog));
                    File.AppendAllText(fallbackLog, $"{DateTime.Now} | CRITICAL STARTUP ERROR: {ex.Message}{Environment.NewLine}");
                }
                catch { }
                Stop();
            }
        }

        protected override void OnStop()
        {
            _timer?.Stop();
            _timer?.Dispose();
        }

        // Wczytanie konfiguracji z pliku JSON
        private void LoadConfig()
        {
            if (!File.Exists(_configPath))
                throw new FileNotFoundException($"Brak pliku konfiguracyjnego: {_configPath}");

            string json = File.ReadAllText(_configPath);
            _config = JsonConvert.DeserializeObject<AppConfig>(json);
        }

        // Zapisanie godz State w pliku konfiguracyjnym
        private void SaveConfigState(DateTime newRunDate)
        {
            try
            {
                // Aktualizujemy tylko datę w obiekcie w pamięci
                _config.State.LastRunDate = newRunDate;

                // Zapisujemy cały obiekt z powrotem do pliku JSON
                string json = JsonConvert.SerializeObject(_config, Formatting.Indented);
                File.WriteAllText(_configPath, json);
            }
            catch (Exception ex)
            {
                LogError("CONFIG_SAVE", ex);
            }
        }

        // Zaplanuj następne uruchomienie
        private void ScheduleNextRun()
        {
            try
            {
                // Przeładuj config
                LoadConfig();

                DateTime now = DateTime.Now;
                DateTime nextRun = new DateTime(now.Year, now.Month, now.Day,
                                                _config.Scheduler.TargetHour,
                                                _config.Scheduler.TargetMinute, 0);

                // Jeśli godzina już minęła dzisiaj, ustaw na jutro
                if (now > nextRun)
                {
                    nextRun = nextRun.AddDays(1);
                }

                TimeSpan timeUntilNextRun = nextRun - now;

                // Zabezpieczenie przed ujemnym czasem
                if (timeUntilNextRun.TotalMilliseconds <= 0)
                {
                    timeUntilNextRun = TimeSpan.FromSeconds(10);
                }

                _timer.Interval = timeUntilNextRun.TotalMilliseconds;
                _timer.Start();
            }
            catch (Exception ex)
            {
                LogError("SCHEDULER", ex);
                // W razie błędu spróbuj za minutę
                _timer.Interval = 60000;
                _timer.Start();
            }
        }

        private async void Timer_Elapsed(object source, System.Timers.ElapsedEventArgs e)
        {
            _timer.Stop(); // Zatrzymaj timer, bo AutoReset = false

            if (_isRunning) return;
            _isRunning = true;

            try
            {
                LoadConfig(); // Przeładuj konfigurację

                // Wywołanie głównej metody
                await ImportInvoice();
            }
            catch (Exception ex)
            {
                LogError("SYSTEM", ex);
            }
            finally
            {
                _isRunning = false;
                // Zaplanuj kolejne uruchomienie (np. za 24h)
                ScheduleNextRun();
            }
        }

        private async Task ImportInvoice()
        {
            try
            {
                if (_config == null) return;

                string token = await GetAccessToken();

                string lastRunFormatted = _config.State.LastRunDate.ToString("yyyy-MM-ddTHH:mm:ssZ");

                using (HttpClient http = new HttpClient())
                {
                    http.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

                    // 1. Pobranie nagłówka faktury
                    string jourUrl = $"{_config.Endpoints.D365Url}{_config.Endpoints.JourEntity}?$filter=InvoiceDate gt {lastRunFormatted}";

                    var jourResponse = await http.GetAsync(jourUrl);
                    if (!jourResponse.IsSuccessStatusCode)
                    {
                        LogError("Odata", new Exception($"Błąd podczas pobierania nagłówka faktury: {jourResponse.ReasonPhrase}"));
                    }

                    // 1. Odczyt JSON nagłówka
                    string jourJson = await jourResponse.Content.ReadAsStringAsync();

                    var jourParsed = JsonConvert.DeserializeObject<CustInvoiceJourResponse>(jourJson);

                    if (jourParsed == null || jourParsed.value == null || !jourParsed.value.Any())
                    {
                        return;
                    }

                    var headersDTO = jourParsed.value.Select(x => MapHeader(x)).ToList();

                    foreach (var invoiceDto in headersDTO)
                    {
                        if (string.IsNullOrEmpty(invoiceDto.NumerObcy))
                            continue;

                        // 2. Pobranie pozycji faktury
                        string transUrl = $"{_config.Endpoints.D365Url}{_config.Endpoints.TransEntity}?$filter=InvoiceId eq '{invoiceDto.NumerObcy}'";

                        var transResponse = await http.GetAsync(transUrl);
                        if (!transResponse.IsSuccessStatusCode)
                        {
                            LogError(invoiceDto.NumerObcy, new Exception($"Błąd podczas pobierania pozycji faktury: {invoiceDto.NumerObcy}"));
                            continue;
                        }

                        string transJson = await transResponse.Content.ReadAsStringAsync();
                        var transParsed = JsonConvert.DeserializeObject<CustInvoiceTransResponse>(transJson);

                        if (transParsed != null && transParsed.value != null)
                        {
                            invoiceDto.Elementy = transParsed.value.Select((x, i) => MapLine(x, i + 1)).ToList();

                            var jsonSettings = new JsonSerializerSettings
                            {
                                Formatting = Formatting.Indented,
                                DateFormatString = "yyyy-MM-ddTHH:mm:ss"
                            };

                            string finalJson = JsonConvert.SerializeObject(invoiceDto, jsonSettings);

                            // Użycie ścieżki z konfiguracji
                            Directory.CreateDirectory(_config.Paths.ExportFolder);
                            string filePath = Path.Combine(_config.Paths.ExportFolder, $"Invoice_{invoiceDto.NumerObcy.Replace("/", "_")}.json");

                            File.WriteAllText(filePath, finalJson);
                        }
                    }
                    // Aktualizacja daty ostatniego uruchomienia
                    SaveConfigState(DateTime.UtcNow);
                }
            }
            catch (Exception e)
            {
                LogError("SYSTEM_IMPORT", e);
            }

        }

        private async Task<string> GetAccessToken()
        {
            var app = ConfidentialClientApplicationBuilder.Create(_config.AzureAd.ClientId)
                .WithClientSecret(_config.AzureAd.ClientSecret)
                .WithAuthority($"https://login.microsoftonline.com/{_config.AzureAd.TenantId}")
                .Build();

            string[] scopes = new[] { $"{_config.Endpoints.D365Url}/.default" };

            AuthenticationResult token = await app.AcquireTokenForClient(scopes).ExecuteAsync();
            return token.AccessToken;
        }

        private void LogError(string invoiceNumer, Exception e)
        {
            string logPath = _config?.Paths?.ErrorLog ?? @"C:\Fallback_log_services\ImportDynamics_error_fallback.txt";
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(logPath));
                File.AppendAllText(logPath, $"{DateTime.Now} | Faktura: {invoiceNumer} | Błąd: {e.Message}{Environment.NewLine}");
            }
            catch { }
        }

        // Mapowanie pol OData -> pola JSON Proferis
        private InvoiceDTO MapHeader(CustInvoiceJourEntity src)
        {
            DateTime invoiceDate = DateTime.Parse(src.InvoiceDate);
            DateTime dueDate = DateTime.Parse(src.DueDate);

            // Pobieranie kodu i kursów
            string supplierCode = LoadOptimaSupplierCode(cleanToDigits(src.VATNum));
            var exchangeRate = LoadExchangeRate(invoiceDate.Date);

            return new InvoiceDTO
            {
                NumerObcy = src.InvoiceId,
                Podmiot = new InvoicePodmiot { Kod = supplierCode },

                DataWystawienia = invoiceDate,
                DataOperacji = invoiceDate,
                TerminPlatnosci = dueDate,
                Waluta = src.CurrencyCode,
                Korekta = IsCorrectionInvoice(src.InvoiceId),

                KursNumer = exchangeRate.KursID,
                DataKursu = exchangeRate.kursData,
                KursL = src.CurrencyCode == "PLN" ? 1 : exchangeRate.KursL,
                KursM = src.CurrencyCode == "PLN" ? 1 : exchangeRate.KursM,

                RodzajTransakcji = src.CustGroup == "OKRAJ" ? "Krajowa" : "Wewnatrzunijna",

                // Inicjalizacja pustej listy elementów
                Elementy = new List<InvoiceLineDTO>()
            };
        }

        private InvoiceLineDTO MapLine(CustInvoiceLineEntity src, int index)
        {
            var TaxRate = Decimal.Parse(src.TaxWriteCode.Trim(new char[] { '%' }));
            string SalesUnit = cleanSalesUnits(src.SalesUnit);

            return new InvoiceLineDTO
            {
                // index
                Pozycja = index,

                Towar = new InvoiceTowar { Kod = src.ItemId },

                Ilosc = src.Qty,
                Cena = src.SalesPrice,
                Stawka = TaxRate,
                WartoscNetto = src.LineAmount,
                WartoscBrutto = src.LineAmount + src.TaxAmount,
                Opis = src.Name,
                JM = SalesUnit
            };
        }

        // Pobranie Kodu kontrahenta z Optima na podstawie numeru NIP
        private string LoadOptimaSupplierCode(string nip)
        {
            using (var con = new SqlConnection(_config.ConnectionStrings.Optima))
            {
                con.Open();

                using (var checkCmd = new SqlCommand(
                    @"SELECT COUNT(*) FROM [CDN_KFM].[CDN].[Kontrahenci]", con))
                {
                    int count = (int)checkCmd.ExecuteScalar();

                    if (count == 0)
                        throw new Exception("Brak kontrahentów w bazie Optima.");
                }

                using (var cmd = new SqlCommand(@"
                    SELECT TOP 1 Knt_Kod
                    FROM [CDN_KFM].[CDN].[Kontrahenci]
                    WHERE Knt_Nip = @NIP", con))
                {
                    cmd.Parameters.AddWithValue("@NIP", nip);

                    object result = cmd.ExecuteScalar();
                    return (result == null || result is DBNull) ? null : result.ToString();
                }
            }
        }

        // Pobranie kursu waluty z OptimaConfig
        private (int KursID, DateTime kursData, decimal KursL, int KursM) LoadExchangeRate(DateTime date)
        {
            using (var con = new SqlConnection(_config.ConnectionStrings.OptimaConfig))
            {
                con.Open();

                using (var cmd = new SqlCommand(@"
                    SELECT TOP (1) WKu_WKuID, WNo_Publikacja, WNo_KursL, WNo_KursM
                    FROM [CDN_KNF_Konfiguracja].[CDN].[WalNotowania] WalNot
                    left join [CDN_KNF_Konfiguracja].[CDN].[WalNazwy] walN ON WalN.WNa_WNaID = WalNot.WNo_WNaID
                    left join [CDN_KNF_Konfiguracja].[CDN].WalKursy walK ON walK.WKu_WKuID = WalNot.WNo_WKuID
                    where WNo_Publikacja < @RateDate
                    and Wna_WNaID = 2 and WKu_WKuID = 3 AND WKu_NieAktywny = 0
                    Order by 2 desc         
                ", con))
                {
                    cmd.Parameters.Add("@RateDate", SqlDbType.Date).Value = date;

                    object result = cmd.ExecuteScalar();

                    using (var r = cmd.ExecuteReader())
                    {
                        if (r.Read())
                        {
                            int kursId = Convert.ToInt32(r["WKu_WKuID"]);
                            DateTime kursData = Convert.ToDateTime(r["WNo_Publikacja"]);
                            decimal kursL = Convert.ToDecimal(r["WNo_KursL"]);
                            int kursM = Convert.ToInt32(r["WNo_KursM"]);

                            return (kursId, kursData, kursL, kursM);
                        }
                        else
                        {
                            throw new Exception($"Brak kursu waluty dla daty {date.ToShortDateString()} w bazie Optima.");
                        }
                    }
                }
            }
        }

        // Clean dignital 
        private string cleanToDigits(string s) => string.IsNullOrEmpty(s) ? "" : new string(s.Where(c => char.IsDigit(c)).ToArray());

        // Covert UnitSales DE to PL
        private string cleanSalesUnits(string s)
        {
            if (string.IsNullOrEmpty(s))
                return "";

            switch (s.ToLowerInvariant())
            {
                case "stk": return "szt";
                case "kg": return "kg";
                case "m": return "m";
                case "lf": return "m";
                case "l": return "l";
                case "m2": return "m2";
                default: return "szt";
            }
        }

        private bool IsCorrectionInvoice(string s)
        {
            if (string.IsNullOrWhiteSpace(s))
                return false;

            s = s.Trim();

            int idx = s.IndexOf('/');
            string firstPart = idx >= 0 ? s.Substring(0, idx) : s;

            switch(firstPart.ToLowerInvariant())
            {
                case "fnk":
                case "fnz":
                case "fs":
                case "rekl-k":
                case "rg":
                    return false;
                case "fnzk":
                case "fsk":
                case "fnkk":
                case "rgk":
                    return true;
                default:
                    return false;
            }
        }
    }
}
