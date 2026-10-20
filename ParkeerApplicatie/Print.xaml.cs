using ParkeerApplicatie.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;

namespace ParkeerApplicatie;

public partial class Print : Window
{
    private static readonly CultureInfo Nl = new("nl-BE");

    // Sleutel is DayOfWeek zelf: geen tikfouten in de dagnaam mogelijk.
    private static readonly Dictionary<DayOfWeek, string> DagKleuren = new()
    {
        [DayOfWeek.Monday] = "Coral",
        [DayOfWeek.Tuesday] = "Blue",
        [DayOfWeek.Wednesday] = "Red",
        [DayOfWeek.Thursday] = "Orange",
        [DayOfWeek.Friday] = "Purple",
        [DayOfWeek.Saturday] = "Yellow",
        [DayOfWeek.Sunday] = "Green",
    };

    // Properties waar de XAML aan bindt.
    public Registratie Registratie { get; }
    public string ColorToday { get; }
    public string DatumVoluit { get; }
    public string TijdTekst => Registratie.TijdInrijden.ToString("HH:mm");
    public string NaamChauffeur => $"{Registratie.Voornaam} {Registratie.Familienaam}".Trim();

    public Print(Registratie registratie)
    {
        InitializeComponent();
        Registratie = registratie;

        // Kleur hoort bij de dag waarvoor de kaart geldt (niet bij het moment van afdrukken).
        ColorToday = DagKleuren[registratie.Datum.DayOfWeek];

        bool termijn = !registratie.DagKaart && registratie.termijnKaart == true;

        DatumVoluit = termijn && registratie.DatumTot is { } tot
            ? $"{Voluit(registratie.Datum)} tot {Voluit(tot)} -"
            : $"{Voluit(registratie.Datum)} -";

        ViewDagKaart.Visibility = registratie.DagKaart ? Visibility.Visible : Visibility.Collapsed;
        ViewtermijnKaart.Visibility = termijn ? Visibility.Visible : Visibility.Collapsed;
        ViewUurKaart.Visibility = !registratie.DagKaart && !termijn ? Visibility.Visible : Visibility.Collapsed;

        // Pas hier, nadat alle properties een waarde hebben: ze melden geen wijzigingen
        // (geen INotifyPropertyChanged), dus een binding die te vroeg leest, blijft leeg.
        DataContext = this;
    }

    private static string Voluit(DateTime d) => d.ToString("dddd d MMMM yyyy", Nl);

    private void Afdrukken_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new PrintDialog();
        if (dialog.ShowDialog() != true) return;

        // De kaart schaalt mee naar het afdrukbare gedeelte van het papier.
        var formaat = new Size(dialog.PrintableAreaWidth, dialog.PrintableAreaHeight);
        Kaart.Measure(formaat);
        Kaart.Arrange(new Rect(formaat));

        dialog.PrintVisual(Kaart, "Parkeerkaart");

        Kaart.InvalidateMeasure(); // terug naar de grootte van het venster
    }
}