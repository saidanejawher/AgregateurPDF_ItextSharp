using iTextSharp.text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Upsideo.Agreagateur.Services
{
    public enum Template { V1, V2 }
    public enum Civilite { M, MME }
    public enum Profil { TresPrudent, Prudent, Equilibre, Dynamique }
    public class Avoir
    {
        public string Titulaire;
        public string Numero;
        public DateTime DateOuverture;
        public string Etablissement;
        public string Nom;
        public string Type;
        public double Valorisation;
        public double PerfYTD;
        public double PerfOrigine;
        public double PerfPeriode;
        public double PerfNMoins1;
        public double PerfNMoins2;
        public double PerfNMoins3;
        public double PMV;
        public Profil Profil;
        public double SommeVersements { get { return Mouvements.Where(x => x.Type == TypeMouvement.Versement).Sum(x => x.Montant); } }
        public double SommeRetraits { get { return Mouvements.Where(x => x.Type == TypeMouvement.Retrait).Sum(x => x.Montant); } }
        public double SRRIMoyen { get { return Positions.Average(x => x.SRRI * x.Poids); } }
       //New code here !!
        public int SommeVersementsMois(int mois)
        {
            Random rnd = new Random();
            return rnd.Next(1, 40);
        }
        public int SommeRetraitsMois(int mois)
        {
            Random rnd = new Random();
            return rnd.Next(-40, -1) ;
        }
        public List<Mouvement> Mouvements;
        public List<Position> Positions;
        public Dictionary<DateTime, double> HistoriqueValorisations;
    }
    public class Frais
    {
        public string Type;
        public double Montant;
        public double Pourcentage;
        public DateTime Date;
    }
    public class Position
    {
        public string ClasseActif;
        public string ISIN;
        public string Libelle;
        public double Quantite;
        public double Cours;
        public double Valorisation;
        public double PAM;//Prix d'Achat Moyen
        public double Poids;
        public double PMVEuros;
        public double PMVPourcent;
        public DateTime DateCours;
        public int SRRI;
    }
    public enum TypeMouvement
    {
        Versement,
        Retrait,
        AchatSouscription,
        VenteRachat,
        AcheteVendu,
        Frais,
        Avance
    }
    public enum ZoneGeographique
    {
        AmeriqueNord,
        AmeriqueSud,
        Europe,
        Afrique,
        Asie,
        Oceanie,
        PaysEmergents
    }
    public class Mouvement
    {
        public TypeMouvement Type;
        public double Montant;
        public double Quantite;
        public DateTime DateEffet;
        public string ISIN;
        public string Libelle;
    }
    public class ReleveSituation
    {
        public BaseColor CouleurBase; // in ofi case its bleu ciel  ->  new BaseColor(4, 169, 218);
        public BaseColor CouleurEcriture; // White OR Black Priory
        public Template Template;
        public string PathLogo;
        public Civilite Civilite;
        public string Nom;
        public string Prenom;
        public string RaisonSociale;
        public DateTime PeriodeDebut;
        public DateTime PeriodeFin;
        public string Footer;//exemple : Cabinet TopConseil – 105, rue de la Boétie 75008 Paris<br>RCS : 123456789 – ORIAS : 123456789 – N° de TVA : 123456789<br>Téléphone : 01 12 34 56 78 – Mail : info @topconseil.com
        public string Footer1; // Cabinet TopConseil – 105, rue de la Boétie 75008 Paris
        public string Footer2; // RCS : 123456789 – ORIAS : 123456789 – N° de TVA : 123456789
        public string Footer3; // Téléphone : 01 12 34 56 78 – Mail : info @topconseil.com

        public double GlobalValorisation;
        public double GlobalPerfOrigine;
        public double GlobalPerfYTD;//depuis le 1er janvier de l'année en cours
        public double GlobalTotalVersements;
        public double GlobalTotalRetraits;
        public double GlobalPMV;//plus ou moins value
        public double GlobalPerfPeriode;

        public List<Avoir> Avoirs;

        public Dictionary<ZoneGeographique, double> RepartitionGeographique;

        public double RisqueClient;

        public Dictionary<DateTime, double> GlobalHistoriqueValorisations;

        public List<Frais> Frais;
    }
}
