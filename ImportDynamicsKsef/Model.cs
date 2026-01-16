using Newtonsoft.Json;
using System;
using System.Collections.Generic;

namespace ImportDynamicsKsef
{
    // MODELE ODPOWIEDZI ODATA
    public class CustInvoiceJourResponse
    {
        public List<CustInvoiceJourEntity> value { get; set; }
    }

    public class CustInvoiceTransResponse
    {
        public List<CustInvoiceLineEntity> value { get; set; }
    }

    public class CustInvoiceJourEntity
    {
        public string InvoiceId { get; set; }
        public string CreatedOn { get; set; }
        public string DocumentDate { get; set; }
        public string InvoiceDate { get; set; }
        public string CurrencyCode { get; set; }
        public string DueDate { get; set; }
        public string VATNum { get; set; }
        public string CustGroup { get; set; }
        public decimal InvoiceAmount { get; set; }
    }

    public class CustInvoiceLineEntity
    {
        public decimal Qty { get; set; }
        public decimal SalesPrice { get; set; }
        public string Name { get; set; }
        public string TaxWriteCode { get; set; }
        public decimal LineAmount { get; set; }
        public decimal TaxAmount { get; set; }
        public string ItemId { get; set; }
        public string SalesUnit { get; set; }
    }

    // MODELE DOCELOWEGO JSON

    // 1. Główny obiekt faktury (Root)
    public class InvoiceDTO
    {
        public string Id { get; set; } = null;
        public string Seria { get; set; } = "FS";
        public string NumerDokumentu { get; set; } = null;
        public string NumerObcy { get; set; } = null; // Numer faktury z Dynamics (dokument)
        public bool Bufor { get; set; } = true;

        // ZAGNIEŻDŻONY PODMIOT
        public InvoicePodmiot Podmiot { get; set; }

        public string Platnik { get; set; } = null;
        public string Odbiorca { get; set; } = null;
        public string Kategoria { get; set; } = null;
        public string Magazyn { get; set; } = null;
        public string MagazynDocelowy { get; set; } = null;
        public bool Korekta { get; set; }

        public DateTime DataWystawienia { get; set; }
        public DateTime DataOperacji { get; set; }
        public DateTime TerminPlatnosci { get; set; }

        public string LiczonaOd { get; set; } = "N"; // Netto
        public int Rabat { get; set; } = 0;
        public string Waluta { get; set; }
        public string FormaPlatnosci { get; set; } = "Przelew";

        // Kursy
        public DateTime? DataKursu { get; set; }
        public int? KursNumer { get; set; }
        public decimal? KursL { get; set; }
        public int? KursM { get; set; }

        // Kwoty
        public decimal? WartoscNetto { get; set; }
        public decimal? WartoscBrutto { get; set; }

        public int Zaplacono { get; set; } = 0;
        public bool MPP { get; set; } = false;
        public bool FakturaVATMarza { get; set; } = false;
        public int TypDokumentu { get; set; } = 302;
        public int RodzajDokumentu { get; set; } = 302000;

        public string RodzajTransakcji { get; set; }

        public string Opis { get; set; }

        // LISTA POZYCJI (Wewnątrz głównego obiektu)
        public List<InvoiceLineDTO> Elementy { get; set; }
    }

    // 2. Struktura Podmiotu
    public class InvoicePodmiot
    {
        public string Id { get; set; } = null;
        public string Kod { get; set; }
        public object Adres { get; set; } = new object();
    }

    // 3. Struktura Pozycji
    public class InvoiceLineDTO
    {
        public string Id { get; set; } = null;
        public int Pozycja { get; set; } // Auto-increment

        // ZAGNIEŻDŻONY TOWAR
        public InvoiceTowar Towar { get; set; }

        public decimal Ilosc { get; set; }
        public string JM { get; set; }
        public decimal Cena { get; set; }
        public decimal Stawka { get; set; }
        public string Flaga { get; set; } = null;
        public decimal WartoscNetto { get; set; }
        public decimal WartoscBrutto { get; set; }
        public int Rabat { get; set; } = 0;
        public string Opis { get; set; }
    }

    // 4. Struktura Towaru
    public class InvoiceTowar
    {
        public int Id { get; set; } = 0;
        public string Kod { get; set; } // Tu ItemId
    }
}