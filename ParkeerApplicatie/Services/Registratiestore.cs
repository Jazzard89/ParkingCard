using ClosedXML.Excel;
using ParkeerApplicatie.Models;
using System.IO;

namespace ParkeerApplicatie.Services;

/// <summary>
/// Leest en schrijft registraties in het Excel-bestand (eerste werkblad).
/// Kolommen worden gevonden via hun koptekst, niet via hun positie. Ontbrekende kolommen worden bij het schrijven toegevoegd.
/// </summary>
public sealed class RegistratieStore
{
    private static readonly (string Sleutel, string Kop)[] Kolommen =
    {
        ("datum", "Datum"),
        ("firma", "Firma"),
        ("nummerplaat", "Nummerplaat"),
        ("familienaam", "Familienaam"),
        ("voornaam", "Voornaam"),
        ("bestemming", "Bestemming"),
        ("contactpersoon", "Contactpersoon"),
        ("vergunning", "Vergunning"),
        ("dagkaart", "Dagkaart"),
        ("tijdinrijden", "Tijd inrijden"),
        ("tijduitrijden", "Tijd uitrijden"),
        ("gsm", "GSM"),
        ("typevoertuig", "Type voertuig"),
    };

    // Andere schrijfwijzen van dezelfde kolom (na normaliseren: kleine letters, zonder spaties en leestekens).
    private static readonly Dictionary<string, string> Aliassen = new()
    {
        ["numerplaat"] = "nummerplaat",
        ["kenteken"] = "nummerplaat",
        ["locatie"] = "bestemming",
        ["zone"] = "bestemming",
        ["naam"] = "familienaam",
        ["dagkaartjanee"] = "dagkaart",
        ["gsmnummer"] = "gsm",
        ["telefoon"] = "gsm",
    };

    public List<Registratie> Laad(string pad)
    {
        using var wb = new XLWorkbook(pad);
        var ws = wb.Worksheet(1);
        var k = LeesKoppen(ws);

        foreach (var verplicht in new[] { "datum", "tijdinrijden", "voornaam", "familienaam" })
        {
            if (!k.ContainsKey(verplicht))
                throw new InvalidDataException($"De kolom '{KopVan(verplicht)}' ontbreekt in het Excel-bestand.");
        }

        string Tekst(IXLRow r, string sleutel) =>
            k.TryGetValue(sleutel, out var c) ? r.Cell(c).GetString().Trim() : "";

        bool Ja(IXLRow r, string sleutel) =>
            Tekst(r, sleutel).Equals("ja", StringComparison.OrdinalIgnoreCase);

        TimeSpan? Tijd(IXLRow r, string sleutel)
        {
            if (!k.TryGetValue(sleutel, out var c)) return null;
            var cel = r.Cell(c);
            if (cel.IsEmpty()) return null;
            if (cel.TryGetValue<TimeSpan>(out var ts)) return ts;
            return TimeSpan.TryParse(cel.GetString(), out ts) ? ts : null;
        }

        var resultaat = new List<Registratie>();
        int laatste = ws.LastRowUsed()?.RowNumber() ?? 1;

        for (int nr = 2; nr <= laatste; nr++)
        {
            var rij = ws.Row(nr);
            if (!rij.Cell(k["datum"]).TryGetValue<DateTime>(out var datum)) continue;
            if (Tijd(rij, "tijdinrijden") is not TimeSpan inrijden) continue;
            var uit = Tijd(rij, "tijduitrijden");

            resultaat.Add(new Registratie
            {
                RijNummer = nr,
                Datum = datum.Date,
                Firma = Tekst(rij, "firma"),
                Nummerplaat = Tekst(rij, "nummerplaat"),
                Familienaam = Tekst(rij, "familienaam"),
                Voornaam = Tekst(rij, "voornaam"),
                Bestemming = Tekst(rij, "bestemming"),
                Contactpersoon = Tekst(rij, "contactpersoon"),
                Vergunning = Ja(rij, "vergunning"),
                DagKaart = Ja(rij, "dagkaart"),
                TijdInrijden = datum.Date + inrijden,
                TijdUitrijden = uit is null ? null : datum.Date + uit.Value,
                Gsm = Tekst(rij, "gsm"),
                TypeVoertuig = Tekst(rij, "typevoertuig"),
            });
        }

        return resultaat;
    }

    /// <summary>Voegt een nieuwe registratie onderaan toe en zet het rijnummer op het object.</summary>
    public void Voeg(string pad, Registratie reg)
    {
        using var wb = new XLWorkbook(pad);
        var ws = wb.Worksheet(1);
        var k = ZorgVoorKoppen(ws);

        int rij = (ws.LastRowUsed()?.RowNumber() ?? 1) + 1;

        ws.Cell(rij, k["datum"]).SetValue(reg.Datum.Date).Style.DateFormat.Format = "dd/MM/yyyy";
        ws.Cell(rij, k["firma"]).SetValue(reg.Firma);
        ws.Cell(rij, k["nummerplaat"]).SetValue(reg.Nummerplaat);
        ws.Cell(rij, k["familienaam"]).SetValue(reg.Familienaam);
        ws.Cell(rij, k["voornaam"]).SetValue(reg.Voornaam);
        ws.Cell(rij, k["bestemming"]).SetValue(reg.Bestemming);
        ws.Cell(rij, k["contactpersoon"]).SetValue(reg.Contactpersoon);
        ws.Cell(rij, k["vergunning"]).SetValue(reg.Vergunning ? "ja" : "nee");
        ws.Cell(rij, k["dagkaart"]).SetValue(reg.DagKaart ? "ja" : "nee");
        ws.Cell(rij, k["tijdinrijden"]).SetValue(reg.TijdInrijden.TimeOfDay).Style.DateFormat.Format = "HH:mm";
        ws.Cell(rij, k["gsm"]).SetValue(reg.Gsm);
        ws.Cell(rij, k["typevoertuig"]).SetValue(reg.TypeVoertuig);

        wb.Save();
        reg.RijNummer = rij;
    }

    /// <summary>Schrijft de uitrijtijd in de rij van deze registratie.</summary>
    public void ZetUitrijden(string pad, Registratie reg, DateTime tijd)
    {
        using var wb = new XLWorkbook(pad);
        var ws = wb.Worksheet(1);
        var k = ZorgVoorKoppen(ws);

        var rij = ws.Row(reg.RijNummer);

        // Controle: staat op die rij nog steeds dezelfde persoon? (bestand kan intussen aangepast zijn)
        bool zelfde =
            rij.Cell(k["datum"]).TryGetValue<DateTime>(out var datum) && datum.Date == reg.Datum.Date
            && rij.Cell(k["familienaam"]).GetString().Trim().Equals(reg.Familienaam, StringComparison.OrdinalIgnoreCase)
            && rij.Cell(k["voornaam"]).GetString().Trim().Equals(reg.Voornaam, StringComparison.OrdinalIgnoreCase);

        if (!zelfde)
            throw new InvalidDataException(
                "De rij in het Excel-bestand komt niet meer overeen met deze registratie " +
                "(is het bestand intussen gewijzigd?). Laad het bestand opnieuw via Bestand > Kies bestand.");

        rij.Cell(k["tijduitrijden"]).SetValue(tijd.TimeOfDay).Style.DateFormat.Format = "HH:mm";
        wb.Save();
    }

    // ---------- hulpmethodes ----------

    private static string Normaliseer(string tekst) =>
        new string(tekst.Where(char.IsLetterOrDigit).ToArray()).ToLowerInvariant();

    private static string KopVan(string sleutel) =>
        Kolommen.First(c => c.Sleutel == sleutel).Kop;

    private static Dictionary<string, int> LeesKoppen(IXLWorksheet ws)
    {
        var map = new Dictionary<string, int>();
        foreach (var cel in ws.Row(1).CellsUsed())
        {
            var sleutel = Normaliseer(cel.GetString());
            if (Aliassen.TryGetValue(sleutel, out var echt)) sleutel = echt;
            map.TryAdd(sleutel, cel.Address.ColumnNumber);
        }
        return map;
    }

    /// <summary>Zoals LeesKoppen, maar voegt ontbrekende kolommen achteraan toe.</summary>
    private static Dictionary<string, int> ZorgVoorKoppen(IXLWorksheet ws)
    {
        var map = LeesKoppen(ws);
        int volgende = map.Count == 0 ? 1 : map.Values.Max() + 1;

        foreach (var (sleutel, kop) in Kolommen)
        {
            if (map.ContainsKey(sleutel)) continue;
            var cel = ws.Cell(1, volgende);
            cel.SetValue(kop);
            cel.Style.Font.Bold = true;
            map[sleutel] = volgende++;
        }
        return map;
    }
}