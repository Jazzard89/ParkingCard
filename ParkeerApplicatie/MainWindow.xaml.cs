using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Win32;
using ParkeerApplicatie.Models;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using Path = System.IO.Path;

namespace ParkeerApplicatie
{
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;
        private readonly ObservableCollection<Registratie> _registraties = new();
        private readonly ObservableCollection<Registratie> _registratiesVandaag = new();

        private ObservableCollection<string> _nameSuggestions = new();
        private ObservableCollection<string> _firstNameSuggestions = new();
        public ObservableCollection<string> NameSuggestions
        {
            get { return _nameSuggestions; }
            set
            {
                _nameSuggestions = value;
                OnPropertyChanged();
            }
        }
        public ObservableCollection<string> FirstNameSuggestions
        {
            get { return _firstNameSuggestions; }
            set
            {
                _firstNameSuggestions = value;
                OnPropertyChanged();
            }
        }

        IOrderedEnumerable<Registratie>? Namen = null;


        private static readonly string VoorkeurenPad = Path.Combine(
    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
    "Voorkeuren.txt");

        public MainWindow()
        {
            InitializeComponent();
            DispatcherTimer tikTak = new DispatcherTimer();
            tikTak.Tick += new EventHandler(dispatcherTimer_Tick);
            tikTak.Interval = new TimeSpan(0, 0, 5);
            tikTak.Start();

            TxtNaam.ItemsSource = NameSuggestions;
            TxtVoornaam.ItemsSource = FirstNameSuggestions;

            GetPreferences();
            DgRegistraties.ItemsSource = _registraties;
            DgRegistraties2.ItemsSource = _registratiesVandaag;
            ZetHuidigeDatumEnTijd();
        }

        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private void dispatcherTimer_Tick(object sender, EventArgs e)
        {
            DgRegistraties2.Items.Refresh();
        }

        #region Preferences
        private void GetPreferences()
        {
            if (!File.Exists(VoorkeurenPad)) return;

            string[] delen = File.ReadAllText(VoorkeurenPad).Split(';');
            if (delen.Length != 2) return;

            string pad = Path.Combine(delen[0], delen[1]);
            if (File.Exists(pad))
            {
                LaadExcel(pad);
            }
        }
        private void SetPreferences(string directory, string filename)
        {
            string allText = $"{directory};{filename}";
            File.WriteAllText(VoorkeurenPad, allText);
        }
        #endregion

        private void ZetHuidigeDatumEnTijd()
        {
            DpDatum.SelectedDate = DateTime.Today;
            TxtTijd.Text = DateTime.Now.ToString("HH:mm");
        }
        private void HuidigeTijd_Click(object sender, RoutedEventArgs e)
        {
            TxtTijd.Text = DateTime.Now.ToString("HH:mm");
        }
        private bool LaadExcel(string pad)
        {
            try
            {
                using var workbook = new XLWorkbook(pad);
                var blad = workbook.Worksheet(1);

                _registraties.Clear();
                _registratiesVandaag.Clear();


                foreach (var rij in blad.RowsUsed().Skip(1))
                {
                    if (!rij.Cell(1).TryGetValue<DateTime>(out DateTime datum)) continue;

                    TimeSpan inrijden = rij.Cell(10).GetTimeSpan();
                    string uitText = rij.Cell(11).GetString();

                    string voornaam = rij.Cell(5).GetString().Trim();
                    string familienaam = rij.Cell(4).GetString().Trim();

                    _registraties.Add(new Registratie
                    {
                        Datum = datum.Date,
                        Firma = rij.Cell(2).GetString().Trim(),
                        Nummerplaat = rij.Cell(3).GetString().Trim(),
                        Familienaam = familienaam,
                        Voornaam = voornaam,
                        Bestemming = rij.Cell(6).GetString().Trim(),
                        GsmContactpersoon = rij.Cell(7).GetString().Trim(),
                        Vergunning = IsJa(rij.Cell(8).GetString()),
                        DagKaart = IsJa(rij.Cell(9).GetString()),
                        TijdInrijden = datum.Date + inrijden,
                        TijdUitrijden = string.IsNullOrWhiteSpace(uitText)
                                            ? null
                                            : datum.Date + rij.Cell(11).GetTimeSpan(),
                        GsmChauffeur = rij.Cell(12).GetString().Trim(),
                        TypeVoertuig = rij.Cell(13).GetString().Trim()
                    });



                    if (datum.Date == DateTime.Today)
                    {
                        _registratiesVandaag.Add(new Registratie
                        {
                            Datum = datum.Date,
                            Firma = rij.Cell(2).GetString().Trim(),
                            Nummerplaat = rij.Cell(3).GetString().Trim(),
                            Familienaam = familienaam,
                            Voornaam = voornaam,
                            Bestemming = rij.Cell(6).GetString().Trim(),
                            GsmContactpersoon = rij.Cell(7).GetString().Trim(),
                            Vergunning = IsJa(rij.Cell(8).GetString()),
                            DagKaart = IsJa(rij.Cell(9).GetString()),
                            TijdInrijden = datum.Date + inrijden,
                            TijdUitrijden = string.IsNullOrWhiteSpace(uitText) ? null : datum.Date + rij.Cell(11).GetTimeSpan(),
                            GsmChauffeur = rij.Cell(12).GetString().Trim(),
                            TypeVoertuig = rij.Cell(13).GetString().Trim()
                        });
                    }

                }
                VulSuggesties();

                return true;
            }
            catch (IOException)
            {
                MessageBox.Show("Het bestand is open in Excel. Sluit het eerst en probeer opnieuw.",
                                "Kan bestand niet lezen", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }



        }
        private void VulSuggesties()
        {
            NameSuggestions.Clear();
            Namen = _registraties
                .Select(r => r).Where(n => !string.IsNullOrWhiteSpace(n.Voornaam)
                && !string.IsNullOrWhiteSpace(n.Familienaam))
                .OrderBy(n => n.Voornaam);


            foreach (var naam in Namen)
            {
                NameSuggestions.Add(naam.Familienaam);
                FirstNameSuggestions.Add(naam.Voornaam);
            }
        }
        private static bool IsJa(string tekst) => tekst.Trim().Equals("Ja", StringComparison.OrdinalIgnoreCase);
        private void CheckBox_Checked(object sender, RoutedEventArgs e) 
        {

        }
        private void CheckBox_Unchecked(object sender, RoutedEventArgs e) 
        {
        
        }
        private void Wis_Click(object sender, RoutedEventArgs e)
        {
            WisFormulier();
        }
        private void Opslaan_Click(object sender, RoutedEventArgs e)
        {

            var typeWagen = (CmbTypeWagen.SelectedItem as ComboBoxItem)?.Content?.ToString();
            var locatie = (CbLocatie.SelectedItem as ComboBoxItem)?.Content?.ToString();

            if (string.IsNullOrWhiteSpace(TxtNaam.Text) || string.IsNullOrWhiteSpace(TxtVoornaam.Text)
                || string.IsNullOrWhiteSpace(TxtFirma.Text) || string.IsNullOrWhiteSpace(TxtGsm.Text)
                || string.IsNullOrWhiteSpace(TxtContactpersoon.Text) || string.IsNullOrWhiteSpace(TxtTijd.Text)
                || string.IsNullOrWhiteSpace(TxtNummerplaat.Text)
                || typeWagen is null || locatie is null)
            {
                MessageBox.Show("Alle velden zijn verplicht.", "Ontbrekende gegevens",
                                MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (DpDatum.SelectedDate is not DateTime datum ||
                !TimeSpan.TryParse(TxtTijd.Text, out TimeSpan tijd))
            {
                MessageBox.Show("Geef een geldige datum en tijd in (bv. 08:15).", "Ongeldige invoer",
                                MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string firma = TxtFirma.Text.Trim();
            bool firmaAanwezig = _registratiesVandaag.Any(r =>
                    r.Firma.Equals(firma, StringComparison.OrdinalIgnoreCase) 
                    && r.Bestemming == locatie);

            if (firmaAanwezig)
            {
                var antwoord = MessageBox.Show(
                    $"Er is al een persoon aanwezig van {firma} op {locatie}. Toch registreren?",
                    "Firma al aanwezig", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (antwoord != MessageBoxResult.Yes) return;
            }





            _registratiesVandaag.Add(new Registratie
            {
                Datum = datum.Date,
                Firma = TxtFirma.Text.Trim().ToLower(),
                Nummerplaat = TxtNummerplaat.Text.Trim().ToUpper(),
                Familienaam = TxtNaam.Text.Trim().ToLower(),
                Voornaam = TxtVoornaam.Text.Trim().ToLower(),
                //Bestemming = CbLocatie.SelectedItem.ToString(),
                Bestemming = (CbLocatie.SelectedItem as ComboBoxItem).Content.ToString() ?? "",
                GsmContactpersoon = TxtContactpersoon.Text.Trim().ToUpper(),
                DagKaart = RbDagkaartJa.IsChecked == true,
                TijdInrijden = datum.Date + tijd,
                GsmChauffeur = TxtGsm.Text.Trim(),
                TypeVoertuig = (CmbTypeWagen.SelectedItem as ComboBoxItem).Content.ToString() ?? "",
            });

            WisFormulier();
        }
        private void Laad_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Kies het Excel bestand",
                Filter = "Excel bestanden (*.xlsx)|*.xlsx"
            };
            if (dialog.ShowDialog() != true) return;

            if (LaadExcel(dialog.FileName))
            {
                // pas onthouden als het inlezen gelukt is
                SetPreferences(Path.GetDirectoryName(dialog.FileName)!, dialog.SafeFileName);
                MessageBox.Show($"{_registraties.Count} registraties ingelezen.", "Import klaar",
                                MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }


        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void WisFormulier()
        {
            TxtNaam.Text = "";
            TxtVoornaam.Text = "";
            TxtVoornaam.SelectedIndex = -1;
            TxtNaam.SelectedIndex = -1;
            TxtNummerplaat.Clear();
            TxtFirma.Clear();
            TxtGsm.Clear();
            TxtContactpersoon.Clear();
            CmbTypeWagen.SelectedIndex = -1;
            RbDagkaartNee.IsChecked = true;
            ZetHuidigeDatumEnTijd();
            TxtNaam.Focus();
        }


        private void TxtNaam_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (TxtNaam.SelectedItem is not string gekozenNaam) return;

            var voornamen = _registraties
                .Where(r => r.Familienaam.Equals(gekozenNaam, StringComparison.OrdinalIgnoreCase))
                .Select(r => r.Voornaam)
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(v => v, StringComparer.OrdinalIgnoreCase)
                .ToList();

            FirstNameSuggestions.Clear();
            foreach (var v in voornamen)
                FirstNameSuggestions.Add(v);
        }

        private void TxtVoornaam_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (TxtNaam.SelectedItem != null || TxtNaam.SelectedIndex != -1 
                && TxtVoornaam.SelectedItem != null || TxtVoornaam.SelectedIndex != -1)
            {
                EnableFields(true);
            }
            else
            {
                EnableFields(false);
            }
        }

        private void Naamvelden_TextChanged(object sender, TextChangedEventArgs e)
        {
            bool beideIngevuld = !string.IsNullOrWhiteSpace(TxtNaam.Text)
                              && !string.IsNullOrWhiteSpace(TxtVoornaam.Text);
            if (beideIngevuld)
            {
                EnableFields(beideIngevuld);

                var vorige = _registraties.LastOrDefault(x =>
                    x.Voornaam.Equals(TxtVoornaam.Text, StringComparison.OrdinalIgnoreCase) &&
                    x.Familienaam.Equals(TxtNaam.Text, StringComparison.OrdinalIgnoreCase));


                if (vorige == null) return;

                TxtNummerplaat.Text = vorige.Nummerplaat;
                TxtFirma.Text = vorige.Firma;
                TxtGsm.Text = vorige.GsmChauffeur;
                TxtContactpersoon.Text = vorige.GsmContactpersoon;

                CmbTypeWagen.SelectedItem = CmbTypeWagen.Items
                    .OfType<ComboBoxItem>()
                    .FirstOrDefault(i => i.Content?.ToString() == vorige.TypeVoertuig);

                CbLocatie.SelectedItem = CbLocatie.Items
                    .OfType<ComboBoxItem>()
                    .FirstOrDefault(i => i.Content?.ToString() == vorige.Bestemming);
            }
        }

        private void EnableFields(bool value)
        {
            TxtNummerplaat.IsEnabled = value;
            CmbTypeWagen.IsEnabled = value;
            TxtFirma.IsEnabled = value;
            TxtGsm.IsEnabled = value;
            TxtContactpersoon.IsEnabled = value;
        }

        private void Row_DoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is not DataGridRow row || row.DataContext is not Registratie reg) return;

            var result = MessageBox.Show($"{reg.Voornaam} {reg.Familienaam} uitschrijven?",
                                         "Uitschrijven", MessageBoxButton.YesNo);
            if (result != MessageBoxResult.Yes) return;

            reg.TijdUitrijden = DateTime.Now;

            _registratiesVandaag.Remove(reg);
        }


        private void OverviewRow_DoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is not DataGridRow row || row.DataContext is not Registratie reg) return;

            var result = MessageBox.Show($"naam: {reg.Voornaam} {reg.Familienaam} | firma: {reg.Firma} | " +
                $"gsm: {reg.GsmChauffeur} contact: {reg.GsmContactpersoon} | nummerplaat: {reg.Nummerplaat} | Dagkaart: {(reg.DagKaart ? "ja" : "nee")}" +
                $"| ingereden: {reg.TijdInrijden} | uitgereden: {reg.TijdUitrijden} | datum: {reg.Datum}",
                                         "Bewerken", MessageBoxButton.YesNo);
            if (result != MessageBoxResult.Yes) return;


            /*
             

             */

            reg.TijdUitrijden = DateTime.Now;

            _registratiesVandaag.Remove(reg);
        }

        private void TxtVoornaam_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
    }
}