namespace ParkeerApplicatie.Models
{
    public class Registratie
    {

        public bool FinalWarning => !DagKaart && 
            TijdInrijden < DateTime.Now.AddMinutes(-75);
        public bool FirstWarning => !DagKaart &&
            TijdInrijden < DateTime.Now.AddHours(-1)
            && TijdInrijden > DateTime.Now.AddMinutes(-75);

        public DateTime Datum { get; set; }
        public string Firma { get; set; } = "";
        public string Nummerplaat { get; set; } = "";
        public string Familienaam { get; set; } = "";
        public string Voornaam { get; set; } = "";
        public string Bestemming { get; set; } = "";
        public string GsmContactpersoon { get; set; } = "";
        public bool Vergunning { get; set; }
        public bool DagKaart { get; set; }
        public DateTime TijdInrijden { get; set; }
        public DateTime? TijdUitrijden { get; set; }
        public string GsmChauffeur { get; set; } = "";
        public string TypeVoertuig { get; set; } = "";
    }
}
