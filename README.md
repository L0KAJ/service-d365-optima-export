# 🧾 OData Invoice Export Service

Windows Service do automatycznego, cyklicznego pobierania faktur z Microsoft Dynamics 365 (OData) i eksportowania ich do formatu JSON, z dodatkowym wzbogacaniem danych z systemu Comarch ERP Optima.

**Środowisko:**

- **Serwer:** `192.168.30.34` (KOLWTA-IIS01)
- **Nazwa usługi systemowej:** `ImportFromDynamics`

---

## 📌 Opis projektu

**OData Invoice Export Service** to usługa systemowa (Windows Service) napisana w C#, której głównymi zadaniami są:

1. **Komunikacja z Dynamics 365:** Automatyczne logowanie (Azure AD) i pobieranie faktur przez OData API.
2. **Pobieranie przyrostowe:** Pobieranie tylko faktur wystawionych po dacie ostatniego sukcesu (`LastRunDate`).
3. **Integracja z SQL (Optima):**
   - Mapowanie NIP (`VATNum`) na Kod Kontrahenta (`Knt_Kod`).
   - Pobieranie kursów walut z tabeli konfiguracyjnej.
4. **Mapowanie i Eksport:** Tworzenie plików JSON o złożonej strukturze (zagnieżdżone obiekty `Podmiot`, `Towar`).
5. **Zarządzanie stanem:** Automatyczna aktualizacja daty ostatniego uruchomienia w pliku konfiguracyjnym.

---

## ⚙️ Funkcjonalności

- ✅ **Pełna parametryzacja:** Konfiguracja przez zewnętrzny plik `appsettings.json`.
- ✅ **Azure AD Auth:** Uwierzytelnianie _Client Credentials Flow_ (MSAL).
- ✅ **OData v4:** Filtrowanie dynamiczne (`InvoiceDate gt LastRunDate`).
- ✅ **SQL Integration:** Połączenie z dwiema bazami Optima (Handlowa i Konfiguracyjna).
- ✅ **Scheduler:** Harmonogram dzienny (np. uruchomienie codziennie o 22:00).
- ✅ **Persystencja:** Zapamiętywanie momentu ostatniego importu.
- ✅ **Logowanie:** Zapisywanie błędów do plików tekstowych.

---

## ⚠️ Wymagania uruchomieniowe (Konto usługi)

Aby usługa działała poprawnie (szczególnie w kontekście dostępu do sieciowych ścieżek zapisu lub specyficznych uprawnień SQL), **musi być uruchomiona na dedykowanym koncie użytkownika**, a nie jako `LocalSystem`.

**Konfiguracja w `services.msc`:**

1. Kliknij prawym przyciskiem na usługę `ImportFromDynamics`.
2. Wybierz **Właściwości** -> zakładka **Logowanie** (Log On).
3. Wybierz opcję **To konto** (This account).
4. Wprowadź poświadczenia użytkownika domenowego/lokalnego (np. `Admin`), który posiada uprawnienia do:
   - Odczytu pliku konfiguracyjnego.
   - Zapisu w folderze eksportu (zwłaszcza jeśli jest to zasób sieciowy).

---

## 📄 Konfiguracja (ImportDynamicsSettings.json)

Plik konfiguracyjny znajduje się w lokalizacji:  
`C:\SettingsServicesOptima\ImportDynamicsSettings.json`

**Przykładowa struktura:**

```json
{
  "AzureAd": {
    "TenantId": "TENANT-ID",
    "ClientId": "CLIENT-ID",
    "ClientSecret": "SECRET-KEY"
  },
  "Endpoints": {
    "D365Url": "https://##########################",
    "JourEntity": "/data/CustInvoiceJourBiEntities",
    "TransEntity": "/data/CustInvoiceTransBiEntities"
  },
  "ConnectionStrings": {
    "Optima": "Server=###\\###;Database=########;User Id=...;Password=...;",
    "OptimaConfig": "Server=###\\###;Database=##########;User Id=...;Password=...;"
  },
  "Scheduler": {
    "TargetHour": 22,
    "TargetMinute": 0
  },
  "State": {
    "LastRunDate": "2026-01-01T00:00:00Z"
  },
  "Paths": {
    "ExportFolder": "C:\\InvoicesOptima\\ImportDynamicsInvoices\\Invoices",
    "ErrorLog": "C:\\InvoicesOptima\\ImportDynamicsInvoices\\Error\\error_log.txt"
  }
}
```

### Kluczowe parametry

- **Scheduler:** Określa godzinę codziennego startu.
- **State.LastRunDate:** Data (UTC) ostatniego poprawnego importu. Jest ona nadpisywana przez usługę po każdym sukcesie.
- **Paths:** Ścieżki lokalne lub sieciowe (wymagają odpowiedniego konta usługi).

---

## 🕒 Harmonogram działania

Usługa działa w trybie **Daily Job**:

1. Po starcie (`OnStart`) wczytuje konfigurację.
2. Oblicza czas do najbliższej godziny `TargetHour:TargetMinute`.
3. Usypia wątek (ustawia Timer) na ten okres.
4. O wyznaczonej godzinie wykonuje import.
5. Po zakończeniu aktualizuje `LastRunDate` i planuje kolejne uruchomienie za 24h.

**Resetowanie pobierania:** Aby pobrać faktury ponownie, zatrzymaj usługę, cofnij datę `LastRunDate` w pliku JSON i uruchom usługę ponownie.

---

## 🔗 Integracja SQL (Optima)

Dane uzupełniane z SQL Server:

- **Kod Kontrahenta:** Pobierany z `[###].[######]` na podstawie NIP (`VATNum` z Dynamics).
- **Kurs Waluty:** Pobierany z `[###].[#######]` (tabela `WalKursy`) na podstawie daty faktury.

---

## 🧩 Struktura wyjściowa JSON

**Format pliku:** `Invoice_{NumerFaktury}.json`

### Przykładowa struktura

```json
{
  "Id": null,
  "Seria": "FS",
  "NumerObcy": "F/123/2026",
  "Podmiot": {
    "Kod": "KOD_Z_OPTIMY"
  },
  "DataWystawienia": "2026-01-15T00:00:00",
  "Waluta": "EUR",
  "RodzajTransakcji": "Krajowa",
  "Elementy": [
    {
      "Pozycja": 1,
      "Towar": { "Kod": "ITEM-ID-Z-D365" },
      "Ilosc": 10.0,
      "Cena": 100.0,
      "WartoscNetto": 1000.0
    }
  ]
}
```

---

## 🗂️ Logi i pliki

- **Eksport JSON:** `C:\InvoicesOptima\ImportDynamicsInvoices\Invoices`
- **Logi błędów:** `C:\InvoicesOptima\ImportDynamicsInvoices\Error\error_log.txt`
- **Konfiguracja:** `C:\SettingsServicesOptima\ImportDynamicsSettings.json`
