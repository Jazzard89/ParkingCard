using System.Collections.ObjectModel;
using System.Windows;

namespace ParkeerApplicatie
{
    public class Registratie
    {
        public string Firma { get; set; } = "";
        public string Naam { get; set; } = "";
        public bool DagKaart { get; set; }
        public DateTime TijdInrijden { get; set; }
    }

    public partial class MainWindow : Window
    {
        private readonly ObservableCollection<Registratie> _registraties = new();

        public MainWindow()
        {
            InitializeComponent();
            DgRegistraties.ItemsSource = _registraties;
            ZetHuidigeDatumEnTijd();
        }

        private void ZetHuidigeDatumEnTijd()
        {
            DpDatum.SelectedDate = DateTime.Today;
            TxtTijd.Text = DateTime.Now.ToString("HH:mm");
        }

        private void HuidigeTijd_Click(object sender, RoutedEventArgs e)
        {
            TxtTijd.Text = DateTime.Now.ToString("HH:mm");
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e) { }
        private void CheckBox_Unchecked(object sender, RoutedEventArgs e) { }

        private void Annuleren_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Opslaan_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TxtNaam.Text) || string.IsNullOrWhiteSpace(TxtFirma.Text))
            {
                MessageBox.Show("Naam en firma zijn verplicht.", "Ontbrekende gegevens",
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

            _registraties.Add(new Registratie
            {
                Firma = TxtFirma.Text.Trim(),
                Naam = $"{TxtVoornaam.Text} {TxtNaam.Text}".Trim(),
                DagKaart = ChkDagkaart.IsChecked == true,
                TijdInrijden = datum.Date + tijd
            });

            WisFormulier();
        }

        private void WisFormulier()
        {
            TxtNaam.Clear();
            TxtVoornaam.Clear();
            TxtNummerplaat.Clear();
            TxtFirma.Clear();
            TxtGsm.Clear();
            TxtBestemming.Clear();
            TxtContactpersoon.Clear();
            CmbTypeWagen.SelectedIndex = -1;
            ChkDagkaart.IsChecked = false;
            ZetHuidigeDatumEnTijd();
            TxtNaam.Focus();
        }
    }
}