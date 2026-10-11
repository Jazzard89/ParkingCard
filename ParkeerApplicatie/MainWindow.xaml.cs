using Microsoft.Win32;
using ParkeerApplicatie.Models;
using ParkeerApplicatie.Services;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Threading;

namespace ParkeerApplicatie;

public partial class MainWindow : Window
{
    private static readonly string VoorkeurenPad = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "Voorkeuren.txt");

    private readonly RegistratieStore _store = new();

    // Eén bron van waarheid: alle registraties uit het bestand.
    private readonly ObservableCollection<Registratie> _alle = new();
    // Gefilterde weergave: wie vandaag aanwezig is (nog niet uitgeschreven).
    private readonly ICollectionView _aanwezigVandaag;

    private readonly ObservableCollection<string> _familienamen = new();
    private readonly ObservableCollection<string> _voornamen = new();

    private string? _excelPad;

    public MainWindow()
    {
        InitializeComponent();

        _aanwezigVandaag = new CollectionViewSource { Source = _alle }.View;
        _aanwezigVandaag.Filter = o => o is Registratie r && r.IsAanwezigVandaag;
        _aanwezigVandaag.SortDescriptions.Add(
            new SortDescription(nameof(Registratie.TijdInrijden), ListSortDirection.Ascending));

        TxtNaam.ItemsSource = _familienamen;
        TxtVoornaam.ItemsSource = _voornamen;
        DgRegistraties.ItemsSource = _alle;
        DgRegistraties2.ItemsSource = _aanwezigVandaag;

        // Her-evalueert filter en waarschuwingskleuren (oranje/rood) regelmatig.
        var timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(15) };
        timer.Tick += (_, _) => _aanwezigVandaag.Refresh();
        timer.Start();

        ZetHuidigeDatumEnTijd();
        LaadVoorkeur();
    }

    #region Bestand en voorkeuren

    private void LaadVoorkeur()
    {
        try
        {
            if (!File.Exists(VoorkeurenPad)) return;

            var tekst = File.ReadAllText(VoorkeurenPad).Trim();
            var delen = tekst.Split(';');
            // Oud formaat: "map;bestandsnaam". Nieuw formaat: volledig pad.
            var pad = delen.Length == 2 ? Path.Combine(delen[0], delen[1]) : tekst;

            if (File.Exists(pad)) Laad(pad);
        }
        catch (IOException) { /* voorkeur is niet kritiek */ }
    }

    private static void BewaarVoorkeur(string pad)
    {
        try { File.WriteAllText(VoorkeurenPad, pad); }
        catch (IOException) { }
    }

    private bool Laad(string pad)
    {
        try
        {
            var rijen = _store.Laad(pad);

            _alle.Clear();
            foreach (var r in rijen) _alle.Add(r);

            _excelPad = pad;
            VulSuggesties();
            _aanwezigVandaag.Refresh();
            return true;
        }
        catch (IOException)
        {
            ToonWaarschuwing("Het bestand is open in Excel. Sluit het eerst en probeer opnieuw.", "Kan bestand niet lezen");
        }
        catch (InvalidDataException ex)
        {
            ToonWaarschuwing(ex.Message, "Ongeldig bestand");
        }
        catch (Exception ex)
        {
            ToonWaarschuwing($"Het bestand kon niet gelezen worden: {ex.Message}", "Kan bestand niet lezen");
        }
        return false;
    }

    private void Laad_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "Kies het Excel bestand",
            Filter = "Excel bestanden (*.xlsx)|*.xlsx"
        };
        if (dialog.ShowDialog() != true) return;

        if (Laad(dialog.FileName))
        {
            BewaarVoorkeur(dialog.FileName);
            MessageBox.Show($"{_alle.Count} registraties ingelezen.", "Import klaar",
                            MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }

    private void Close_Click(object sender, RoutedEventArgs e) => Close();

    /// <summary>Voert een schrijfactie op het Excel-bestand uit en toont een duidelijke melding bij problemen.</summary>
    private bool ProbeerSchrijven(Action schrijf)
    {
        try
        {
            schrijf();
            return true;
        }
        catch (IOException)
        {
            ToonWaarschuwing("Het Excel-bestand is open in een ander programma. Sluit het en probeer opnieuw.",
                             "Kan niet opslaan");
        }
        catch (InvalidDataException ex)
        {
            ToonWaarschuwing(ex.Message, "Kan niet opslaan");
        }
        return false;
    }

    private static void ToonWaarschuwing(string tekst, string titel) =>
        MessageBox.Show(tekst, titel, MessageBoxButton.OK, MessageBoxImage.Warning);

    #endregion

    #region Inschrijven

    private void ZetHuidigeDatumEnTijd()
    {
        DpDatum.SelectedDate = DateTime.Today;
        TxtTijd.Text = DateTime.Now.ToString("HH:mm");
    }

    private void HuidigeTijd_Click(object sender, RoutedEventArgs e) =>
        TxtTijd.Text = DateTime.Now.ToString("HH:mm");

    private void Wis_Click(object sender, RoutedEventArgs e) => WisFormulier();

    private void Opslaan_Click(object sender, RoutedEventArgs e)
    {
        if (_excelPad is null)
        {
            ToonWaarschuwing("Kies eerst een Excel-bestand via Bestand > Kies bestand.", "Geen bestand");
            return;
        }

        var naam = TxtNaam.Text.Trim();
        var voornaam = TxtVoornaam.Text.Trim();
        var firma = TxtFirma.Text.Trim();
        var gsm = TxtGsm.Text.Trim();
        var contact = TxtContactpersoon.Text.Trim();
        var plaat = TxtNummerplaat.Text.Trim();
        var typeWagen = (CmbTypeWagen.SelectedItem as ComboBoxItem)?.Content?.ToString();
        var locatie = (CbLocatie.SelectedItem as ComboBoxItem)?.Content?.ToString();

        if (new[] { naam, voornaam, firma, gsm, contact, plaat }.Any(string.IsNullOrWhiteSpace)
            || typeWagen is null || locatie is null)
        {
            ToonWaarschuwing("Alle velden zijn verplicht.", "Ontbrekende gegevens");
            return;
        }

        // TryParseExact: TimeSpan.TryParse("8") zou 8 dagen geven.
        if (DpDatum.SelectedDate is not DateTime datum
            || !TimeSpan.TryParseExact(TxtTijd.Text.Trim(), new[] { @"hh\:mm", @"h\:mm" },
                                       CultureInfo.InvariantCulture, out var tijd))
        {
            ToonWaarschuwing("Geef een geldige datum en tijd in (bv. 08:15).", "Ongeldige invoer");
            return;
        }

        // Waarschuwing: staat er al iemand van dezelfde firma op dezelfde zone? Doorgaan blijft mogelijk.
        var reedsAanwezig = _alle.FirstOrDefault(r =>
            r.IsAanwezigVandaag
            && r.Firma.Equals(firma, StringComparison.OrdinalIgnoreCase)
            && Registratie.ZonesOverlappen(r.Bestemming, locatie));

        if (reedsAanwezig is not null)
        {
            var antwoord = MessageBox.Show(
                $"Er is al een persoon aanwezig van {firma} op {reedsAanwezig.Bestemming} " +
                $"({reedsAanwezig.Voornaam} {reedsAanwezig.Familienaam}).\n\nToch registreren op {locatie}?",
                "Firma al aanwezig", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (antwoord != MessageBoxResult.Yes) return;
        }

        var reg = new Registratie
        {
            Datum = datum.Date,
            Firma = firma,
            Nummerplaat = plaat.ToUpperInvariant(),
            Familienaam = naam,
            Voornaam = voornaam,
            Bestemming = locatie,
            Contactpersoon = contact,
            DagKaart = RbDagkaartJa.IsChecked == true,
            TijdInrijden = datum.Date + tijd,
            Gsm = gsm,
            TypeVoertuig = typeWagen,
        };

        if (!ProbeerSchrijven(() => _store.Voeg(_excelPad, reg))) return;

        _alle.Add(reg);
        VulSuggesties();
        WisFormulier();
    }

    private void WisFormulier()
    {
        TxtNaam.Text = "";
        TxtVoornaam.Text = "";
        TxtNaam.SelectedIndex = -1;
        TxtVoornaam.SelectedIndex = -1;
        TxtNummerplaat.Clear();
        TxtFirma.Clear();
        TxtGsm.Clear();
        TxtContactpersoon.Clear();
        CmbTypeWagen.SelectedIndex = -1;
        CbLocatie.SelectedIndex = -1;
        RbDagkaartNee.IsChecked = true;
        EnableFields(false);
        ZetHuidigeDatumEnTijd();
        TxtNaam.Focus();
    }

    private void EnableFields(bool value)
    {
        TxtNummerplaat.IsEnabled = value;
        CmbTypeWagen.IsEnabled = value;
        TxtFirma.IsEnabled = value;
        TxtGsm.IsEnabled = value;
        TxtContactpersoon.IsEnabled = value;
    }

    #endregion

    #region Suggesties en automatisch aanvullen

    private void VulSuggesties()
    {
        Vul(_familienamen, _alle.Select(r => r.Familienaam));
        Vul(_voornamen, _alle.Select(r => r.Voornaam));
    }

    private static void Vul(ObservableCollection<string> doel, IEnumerable<string> bron)
    {
        doel.Clear();
        foreach (var s in bron
                     .Where(s => !string.IsNullOrWhiteSpace(s))
                     .Distinct(StringComparer.OrdinalIgnoreCase)
                     .OrderBy(s => s, StringComparer.CurrentCultureIgnoreCase))
        {
            doel.Add(s);
        }
    }

    private void TxtNaam_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (TxtNaam.SelectedItem is not string gekozenNaam) return;

        Vul(_voornamen, _alle
            .Where(r => r.Familienaam.Equals(gekozenNaam, StringComparison.OrdinalIgnoreCase))
            .Select(r => r.Voornaam));
    }

    private void Naamvelden_TextChanged(object sender, TextChangedEventArgs e)
    {
        var naam = TxtNaam.Text.Trim();
        var voornaam = TxtVoornaam.Text.Trim();

        var beideIngevuld = naam.Length > 0 && voornaam.Length > 0;
        EnableFields(beideIngevuld);
        if (!beideIngevuld) return;

        // Vul de gegevens van het vorige bezoek van deze persoon alvast in.
        var vorige = _alle.LastOrDefault(x =>
            x.Voornaam.Equals(voornaam, StringComparison.OrdinalIgnoreCase) &&
            x.Familienaam.Equals(naam, StringComparison.OrdinalIgnoreCase));
        if (vorige is null) return;

        TxtNummerplaat.Text = vorige.Nummerplaat;
        TxtFirma.Text = vorige.Firma;
        TxtGsm.Text = vorige.Gsm;
        TxtContactpersoon.Text = vorige.Contactpersoon;

        CmbTypeWagen.SelectedItem = CmbTypeWagen.Items.OfType<ComboBoxItem>()
            .FirstOrDefault(i => i.Content?.ToString() == vorige.TypeVoertuig);
        CbLocatie.SelectedItem = CbLocatie.Items.OfType<ComboBoxItem>()
            .FirstOrDefault(i => i.Content?.ToString() == vorige.Bestemming);
    }

    #endregion

    #region Uitschrijven en overzicht

    private void Row_DoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (sender is not DataGridRow row || row.DataContext is not Registratie reg) return;
        if (_excelPad is null) return;

        var antwoord = MessageBox.Show($"{reg.Voornaam} {reg.Familienaam} uitschrijven?",
                                       "Uitschrijven", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (antwoord != MessageBoxResult.Yes) return;

        var nu = DateTime.Now;
        nu = new DateTime(nu.Year, nu.Month, nu.Day, nu.Hour, nu.Minute, 0);

        if (!ProbeerSchrijven(() => _store.ZetUitrijden(_excelPad, reg, nu))) return;

        reg.TijdUitrijden = nu;
        _aanwezigVandaag.Refresh();      // verdwijnt uit "Momenteel on site"
        DgRegistraties.Items.Refresh();  // toont de uitrijtijd in het overzicht
    }

    private void OverviewRow_DoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (sender is not DataGridRow row || row.DataContext is not Registratie reg) return;

        MessageBox.Show(
            $"Naam: {reg.Voornaam} {reg.Familienaam}\n" +
            $"Firma: {reg.Firma}\n" +
            $"GSM: {reg.Gsm}\n" +
            $"Contactpersoon: {reg.Contactpersoon}\n" +
            $"Nummerplaat: {reg.Nummerplaat} ({reg.TypeVoertuig})\n" +
            $"Locatie: {reg.Bestemming}\n" +
            $"Dagkaart: {(reg.DagKaart ? "ja" : "nee")}\n" +
            $"Datum: {reg.Datum:dd/MM/yyyy}\n" +
            $"Ingereden: {reg.TijdInrijden:HH:mm}\n" +
            $"Uitgereden: {(reg.TijdUitrijden is { } uit ? uit.ToString("HH:mm") : "-")}",
            "Details", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    #endregion
}