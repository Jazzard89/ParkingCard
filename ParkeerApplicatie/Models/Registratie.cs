namespace ParkeerApplicatie.Models;

public class Registratie
{
    // Drempels voor de waarschuwingen (minuten sinds inrijden). Alleen hier aanpassen.
    public const int OranjeNaMinuten = 60;
    public const int RoodNaMinuten = 75;

    public const string ZoneAlle = "Alle";

    /// <summary>Rijnummer in het Excel-bestand, nodig om later de uitrijtijd in dezelfde rij te schrijven.</summary>
    public int RijNummer { get; set; }

    public DateTime Datum { get; set; }
    public DateTime? DatumTot { get; set; }

    public string Firma { get; set; } = "";
    public string Nummerplaat { get; set; } = "";
    public string Familienaam { get; set; } = "";
    public string Voornaam { get; set; } = "";
    public string Bestemming { get; set; } = "";
    public string Contactpersoon { get; set; } = "";
    public bool Vergunning { get; set; }
    public bool DagKaart { get; set; }
    public bool? termijnKaart { get; set; }
    public DateTime TijdInrijden { get; set; }
    public DateTime? TijdUitrijden { get; set; }
    public string Gsm { get; set; } = "";
    public string TypeVoertuig { get; set; } = "";

    public bool IsAanwezig => TijdUitrijden is null;
    public bool IsAanwezigVandaag => IsAanwezig && Datum.Date == DateTime.Today;

    private double MinutenAanwezig => (DateTime.Now - TijdInrijden).TotalMinutes;

    public bool FinalWarning => IsAanwezig && !DagKaart && MinutenAanwezig > RoodNaMinuten;

    public bool FirstWarning => IsAanwezig && !DagKaart
                                && MinutenAanwezig > OranjeNaMinuten
                                && !FinalWarning;

    /// <summary>Twee zones botsen als ze gelijk zijn, of als een van beide "Alle" is.</summary>
    public static bool ZonesOverlappen(string a, string b) =>
        a.Equals(b, StringComparison.OrdinalIgnoreCase)
        || a.Equals(ZoneAlle, StringComparison.OrdinalIgnoreCase)
        || b.Equals(ZoneAlle, StringComparison.OrdinalIgnoreCase);
}