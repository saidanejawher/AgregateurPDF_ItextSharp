using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
//test
using System.Configuration;
using System.Web.UI.DataVisualization.Charting;
using System.Globalization;

namespace Upsideo.Agreagateur.Services
{
    public class BuildPdf
    {
        static string TemplatesDirectory => ConfigurationManager.AppSettings.Get("Directory.Templates");
        static string Directory => ConfigurationManager.AppSettings.Get("Directory");
        static string docReader => ConfigurationManager.AppSettings.Get("docReader");
        static string PathUpImage => ConfigurationManager.AppSettings.Get("PathUpImage");
        static string PathDownImage => ConfigurationManager.AppSettings.Get("PathDownImage");
        static string PathEqualImage => ConfigurationManager.AppSettings.Get("PathEqualImage");

        public static bool BuildAgregateurPDF(ReleveSituation releveSituation, string outFile)
        {
            var reader = new PdfReader(docReader);
            using (var fileStream = new FileStream(Path.Combine(Directory, outFile), FileMode.Create, FileAccess.Write))
            {
                try
                {
                    var document = new Document(reader.GetPageSizeWithRotation(1));
                    var stamper = new PdfStamper(reader, fileStream);
                    document.Open();
                    var ListPathToDelete = new List<string>();
                    var ListSommaire = new List<Tuple<string, int>>();

                    InitObject(releveSituation); // init object evite null reference exception !


                    #region Font ITEXTSHARP
                    var FontColour = releveSituation.CouleurEcriture; //BaseColor.WHITE;
                    Font FontWhite = FontFactory.GetFont("Calibri", 30, FontColour);
                    Font FontBoldWhite = FontFactory.GetFont("Calibri", 18, Font.BOLD, FontColour);
                    Font FontWhitePetite = FontFactory.GetFont("Calibri", 12, FontColour);
                    Font FontWhitTropePetite = FontFactory.GetFont("Calibri", 8, FontColour);
                    Font FontWhitTropePetiteBold = FontFactory.GetFont("Calibri", 8, Font.BOLD, FontColour);
                    Font FontWhitBold_13 = FontFactory.GetFont("Calibri", 13, Font.BOLD, FontColour);
                    Font FontColorSocieteTitle = FontFactory.GetFont("Calibri", 18, releveSituation.CouleurBase);
                    Font FontColorSocieteTitleGrand = FontFactory.GetFont("Calibri", 20, releveSituation.CouleurBase);
                    BaseColor fontColourBlackGris = new BaseColor(65, 64, 66);
                    Font FontSommaire = FontFactory.GetFont("Calibri", 14, fontColourBlackGris);
                    Font FontSimpleEcriture = FontFactory.GetFont("Calibri", 7, fontColourBlackGris);
                    Font FontTitleNoir_16 = FontFactory.GetFont("Calibri", 16, Font.BOLD, fontColourBlackGris);
                    Font FontSimpleEcritureColorSociete = FontFactory.GetFont("Calibri", 7, Font.BOLD, releveSituation.CouleurBase);
                    #endregion
                    var cb = stamper.GetOverContent(1);

                    #region First Page
                    // Add Logo :
                    var PathLogo = releveSituation.PathLogo;
                    Image Logo = Image.GetInstance(PathLogo);
                    Logo.SetAbsolutePosition(PageSize.A4.Width - 200, PageSize.A4.Height - 80);
                    Logo.ScaleAbsoluteHeight(Logo.Height * 0.9f /*40*/);
                    Logo.ScaleAbsoluteWidth(Logo.Width * 0.9f /*190*/);
                    cb.AddImage(Logo);

                    //draw shape 2 : 
                    cb.Rectangle(0, 0, PageSize.A4.Width - 205, 215);

                    cb.MoveTo(PageSize.A4.Width - 205, 0);
                    cb.LineTo(PageSize.A4.Width - 205, 215);
                    cb.LineTo(PageSize.A4.Width - 10, 0);



                    var ColorShape = releveSituation.CouleurBase; //Color of shape
                    cb.SetColorFill(ColorShape);
                    cb.Fill();
                    cb.SetColorStroke(ColorShape);
                    cb.FillStroke();
                    cb.Stroke();
                    cb.ClosePathStroke();


                    var CiviliteNomPrenomClient = $"{releveSituation.Civilite.ToString()}. {releveSituation.Prenom} {releveSituation.Nom}";
                    var PeriodeDebut = $"Du {releveSituation.PeriodeDebut.ToString("dd/MM/yyyy")} ";
                    var PeriodeFin = $"au {releveSituation.PeriodeFin.ToString("dd/MM/yyyy")} ";

                    //draw shape 1 : 
                    cb.Rectangle(0, 0, PageSize.A4.Width - 170, 300);

                    // First trinagle
                    cb.MoveTo(PageSize.A4.Width - 170, 0);
                    cb.LineTo(PageSize.A4.Width + 100, 0);
                    cb.LineTo(PageSize.A4.Width - 170, 300);

                    PdfGState gs1 = new PdfGState();
                    gs1.FillOpacity = 0.5f;
                    cb.SetGState(gs1);

                    cb.SetColorFill(ColorShape);
                    cb.Fill();
                    cb.SetColorStroke(ColorShape);

                    //add RELEVÉ DE SITUATION
                    for (int i = 0; i < 4; i++) // pour enlever l'opacity il faut faire des calques 4x :
                    {
                        ColumnText.ShowTextAligned(cb, Element.ALIGN_LEFT, new Phrase("RELEVÉ DE SITUATION", FontWhite), 40, 230, 0);
                        ColumnText.ShowTextAligned(cb, Element.ALIGN_LEFT, new Phrase(CiviliteNomPrenomClient, FontBoldWhite), 40, 160, 0);
                        ColumnText.ShowTextAligned(cb, Element.ALIGN_LEFT, new Phrase(PeriodeDebut, FontWhitePetite), 40, 130, 0);
                        ColumnText.ShowTextAligned(cb, Element.ALIGN_LEFT, new Phrase(PeriodeFin, FontWhitePetite), 40, 110, 0);

                        ColumnText.ShowTextAligned(cb, Element.ALIGN_LEFT, new Phrase(releveSituation.Footer1, FontWhitTropePetite), 40, 55, 0);
                        ColumnText.ShowTextAligned(cb, Element.ALIGN_LEFT, new Phrase(releveSituation.Footer2, FontWhitTropePetite), 40, 45, 0);
                        ColumnText.ShowTextAligned(cb, Element.ALIGN_LEFT, new Phrase(releveSituation.Footer3, FontWhitTropePetite), 40, 35, 0);
                    }

                    cb.FillStroke();
                    cb.Stroke();
                    cb.ClosePathStroke();

                    document.Close();
                    stamper.Close();
                    #endregion

                    // Now it's for the seconde document !!
                    var TemporaryDoc = "OtherPageTemporary.pdf";

                    using (var Filestream2 = new FileStream(Path.Combine(Directory, TemporaryDoc), FileMode.Create, FileAccess.Write))
                    {
                        var doc1 = new Document();
                        var writer = PdfWriter.GetInstance(doc1, Filestream2);

                        doc1.Open();

                        #region Table Mtiere To Tableau Avoir

                        PdfContentByte cb2 = writer.DirectContent;
                        ColumnText.ShowTextAligned(cb2, Element.ALIGN_LEFT, new Phrase("TABLE DES MATIÈRES", FontColorSocieteTitle), 85, 720, 0);


                        // 3eme Page
                        doc1.NewPage();
                        ColumnText.ShowTextAligned(cb2, Element.ALIGN_LEFT, new Phrase("SITUATION DE VOS AVOIRS", FontColorSocieteTitleGrand), 30, PageSize.A4.Height - 135, 0);
                        ColumnText.ShowTextAligned(cb2, Element.ALIGN_LEFT, new Phrase("Titulaires : ", FontSimpleEcriture), 30, PageSize.A4.Height - 155, 0);

                        var TitulaireDistinct = releveSituation.Avoirs != null ? releveSituation.Avoirs.GroupBy(l => l.Titulaire).Select(g => g.First()).ToList() : new List<Avoir>(); // evite Null reference exception
                        var TitulaireLine = "";
                        string TitulaireLine2 = null;
                        int i = 0;
                        foreach (var a in TitulaireDistinct)
                        {
                            if (i > 10) // Second line
                            {
                                TitulaireLine2 += $" {a.Titulaire} -";
                                i++;
                            }
                            else
                            {
                                TitulaireLine += $" {a.Titulaire} -";
                                i++;
                            }
                        }
                        if (TitulaireLine.Length > 0)
                        {
                            TitulaireLine = TitulaireLine.Substring(0, TitulaireLine.Length - 1);
                            ColumnText.ShowTextAligned(cb2, Element.ALIGN_LEFT, new Phrase(TitulaireLine, FontSimpleEcritureColorSociete), 65, PageSize.A4.Height - 155, 0);
                        }

                        if (TitulaireLine2 != null)
                        {
                            TitulaireLine2 = TitulaireLine2.Substring(0, TitulaireLine2.Length - 1);
                            ColumnText.ShowTextAligned(cb2, Element.ALIGN_LEFT, new Phrase(TitulaireLine2, FontSimpleEcritureColorSociete), 30, PageSize.A4.Height - 175, 0);
                        }

                        var HeightCadreBleu = TitulaireDistinct.Count <= 10 ? PageSize.A4.Height - 230 : PageSize.A4.Height - 250;
                        cb2.Rectangle(30, HeightCadreBleu, 170, 60);
                        cb2.Rectangle(210, HeightCadreBleu, 170, 60);
                        cb2.Rectangle(390, HeightCadreBleu, 170, 60);
                        cb2.SetColorFill(ColorShape);
                        cb2.Fill();
                        cb2.SetColorStroke(ColorShape);
                        cb2.FillStroke();
                        cb2.Stroke();
                        cb2.ClosePathStroke();

                        ColumnText.ShowTextAligned(cb2, Element.ALIGN_LEFT, new Phrase("Valorisation", FontWhitTropePetiteBold), 90, PageSize.A4.Height - 188, 0);
                        ColumnText.ShowTextAligned(cb2, Element.ALIGN_LEFT, new Phrase("Performance depuis l’origine ", FontWhitTropePetiteBold), 240, PageSize.A4.Height - 188, 0);
                        ColumnText.ShowTextAligned(cb2, Element.ALIGN_LEFT, new Phrase("Performance depuis le 1er janvier ", FontWhitTropePetiteBold), 410, PageSize.A4.Height - 188, 0);

                        ColumnText.ShowTextAligned(cb2, Element.ALIGN_LEFT, new Phrase($"{releveSituation.GlobalValorisation.ToString("### ###")}€", FontWhitBold_13), 87, PageSize.A4.Height - 210, 0);
                        var PathGlobalPerfOrigine = releveSituation.GlobalPerfOrigine > 0 ? PathUpImage : releveSituation.GlobalPerfOrigine < 0 ? PathDownImage : PathEqualImage;
                        Image ImageGlobalPerfOrigine = Image.GetInstance(PathGlobalPerfOrigine);
                        ImageGlobalPerfOrigine.SetAbsolutePosition(250, PageSize.A4.Height - 210);
                        ImageGlobalPerfOrigine.ScaleAbsoluteHeight(ImageGlobalPerfOrigine.Height * 0.5f);
                        ImageGlobalPerfOrigine.ScaleAbsoluteWidth(ImageGlobalPerfOrigine.Width * 0.5f);
                        cb2.AddImage(ImageGlobalPerfOrigine);
                        ColumnText.ShowTextAligned(cb2, Element.ALIGN_LEFT, new Phrase($"{releveSituation.GlobalPerfOrigine.ToString("##.#")}%", FontWhitBold_13), 280, PageSize.A4.Height - 210, 0);

                        ColumnText.ShowTextAligned(cb2, Element.ALIGN_LEFT, new Phrase($"{releveSituation.GlobalPerfYTD.ToString("##.#")}%", FontWhitBold_13), 460, PageSize.A4.Height - 210, 0);
                        var PathPerfDepuisJanvier = releveSituation.GlobalPerfYTD > 0 ? PathUpImage : releveSituation.GlobalPerfYTD < 0 ? PathDownImage : PathEqualImage;
                        Image ImagePerfDepuisJanvier = Image.GetInstance(PathPerfDepuisJanvier);
                        ImagePerfDepuisJanvier.SetAbsolutePosition(430, PageSize.A4.Height - 210);
                        ImagePerfDepuisJanvier.ScaleAbsoluteHeight(ImageGlobalPerfOrigine.Height * 0.5f);
                        ImagePerfDepuisJanvier.ScaleAbsoluteWidth(ImageGlobalPerfOrigine.Width * 0.5f);
                        cb2.AddImage(ImagePerfDepuisJanvier);

                        var GrisClair = new BaseColor(247, 247, 247);
                        cb2.Rectangle(30, HeightCadreBleu - 55, PageSize.A4.Width - 65, 45);
                        cb2.SetColorFill(GrisClair);
                        cb2.Fill();
                        cb2.SetColorStroke(GrisClair);
                        cb2.FillStroke();
                        cb2.Stroke();
                        cb2.ClosePathStroke();

                        ColumnText.ShowTextAligned(cb2, Element.ALIGN_LEFT, new Phrase("Somme des versements", FontSimpleEcriture), 70, HeightCadreBleu - 32, 0);
                        ColumnText.ShowTextAligned(cb2, Element.ALIGN_LEFT, new Phrase("Somme des retraits ", FontSimpleEcriture), 190, HeightCadreBleu - 32, 0);
                        ColumnText.ShowTextAligned(cb2, Element.ALIGN_LEFT, new Phrase("+/- value ", FontSimpleEcriture), 350, HeightCadreBleu - 32, 0);
                        ColumnText.ShowTextAligned(cb2, Element.ALIGN_LEFT, new Phrase("Performance entre le 1er janvier", FontSimpleEcriture), 440, HeightCadreBleu - 25, 0);
                        ColumnText.ShowTextAligned(cb2, Element.ALIGN_LEFT, new Phrase($"{releveSituation.PeriodeDebut.ToString("yyyy")} et le {releveSituation.PeriodeDebut.ToString("dd MMM yyyy")}", FontSimpleEcriture), 455, HeightCadreBleu - 32, 0);

                        ColumnText.ShowTextAligned(cb2, Element.ALIGN_LEFT, new Phrase($"{releveSituation.GlobalTotalVersements.ToString("### ###")}€", FontSimpleEcritureColorSociete), 87, HeightCadreBleu - 47, 0);
                        ColumnText.ShowTextAligned(cb2, Element.ALIGN_LEFT, new Phrase($"{releveSituation.GlobalTotalRetraits.ToString("### ###")}€", FontSimpleEcritureColorSociete), 204, HeightCadreBleu - 47, 0);
                        ColumnText.ShowTextAligned(cb2, Element.ALIGN_LEFT, new Phrase($"{releveSituation.GlobalPMV.ToString("### ###")}€", FontSimpleEcritureColorSociete), 355, HeightCadreBleu - 47, 0);
                        ColumnText.ShowTextAligned(cb2, Element.ALIGN_LEFT, new Phrase($"{releveSituation.GlobalPerfPeriode.ToString("##.#")}%", FontSimpleEcritureColorSociete), 475, HeightCadreBleu - 47, 0);
                        var Xseparator = 165;
                        for (var p = 0; p < 3; p++)
                        {
                            ColumnText.ShowTextAligned(cb2, Element.ALIGN_LEFT, new Phrase("|", FontSimpleEcriture), Xseparator, HeightCadreBleu - 35, 0);
                            Xseparator += 130;
                        }
                        #endregion

                        //AVOIRS : 
                        ColumnText.ShowTextAligned(cb2, Element.ALIGN_LEFT, new Phrase("Avoirs", FontTitleNoir_16), 30, HeightCadreBleu - 120, 0);
                        var Couleur_Base_HEXA = System.Drawing.ColorTranslator.ToHtml(System.Drawing.Color.FromArgb(ColorShape.R, ColorShape.G, ColorShape.B));
                        var Couleur_Base_Drawing = System.Drawing.Color.FromArgb(ColorShape.R, ColorShape.G, ColorShape.B);
                        StringBuilder sb = new StringBuilder();
                        string HeaderTable = "<thead> <tr style='height: 39px; color:white; font-family:Calibri;'> <th class='thBackground'>Titulaire</th> <th class='thBackground'>Référence contrat</th> <th class='thBackground'>Date d'ouverture</th> <th class='thBackground'>Etablissement</th> " +
                            "<th class='thBackground'>Type</th> <th class='thBackground'>Valorisation</th> <th class='thBackground'>SRRI*</th> <th class='thBackground'>Perf YTD**</th>" +
                            "<th class='thBackground'>Perf origine</th> </tr> </thead>";

                        #region First Table
                        sb.Append("<html><body><div class='hei'></div><div><table style='width:700px; font-size:12px;  text-align: center; repeat-header: yes; border:0px solid; border-collapse: collapse;'>");
                        sb.Append(HeaderTable);

                        var NbPageVueGlobal = 3;
                        var count = 1;
                        sb.Append("<tbody>");
                        if (releveSituation.Avoirs != null)
                        {
                            foreach (var l in releveSituation.Avoirs)
                            {
                                var CouleurPrefYTD = l.PerfYTD < 0 ? "#f71818" : Couleur_Base_HEXA;
                                var CouleurPrefOrigine = l.PerfOrigine < 0 ? "#f71818" : Couleur_Base_HEXA;
                                if (count % 2 == 0)
                                {
                                    sb.Append("<tr class='tdClass' style='height: 39px; background-color:#efeded; font-family:Calibri;'>");
                                }
                                else
                                {
                                    sb.Append("<tr class='gris tdClass' style='height: 39px; background-color:#f9f9f9; font-family:Calibri;'>");
                                }

                                sb.Append($"<td>{l.Titulaire }</td><td><span style='font-weight: bold;'>{l.Nom} </span><br/><span style='font-size:8;'>{l.Numero} </span></td><td>{l.DateOuverture.ToString("dd/mm/yyy")}</td><td>{l.Etablissement}</td><td>{l.Type}</td><td>{l.Valorisation.ToString("### ###")}</td><td>{l.SRRIMoyen.ToString("#.#")}</td><td style='color:{CouleurPrefYTD}'>{l.PerfYTD.ToString("#.#\\%")}</td><td style='color:{CouleurPrefOrigine}'>{ l.PerfOrigine.ToString("#.#\\%") }</td>");
                                sb.Append("</tr>");
                                if (count % 12 == 0 && count != releveSituation.Avoirs.Count) // la deuxieme condition c'est pour ne pas rajouter un saut de page apres le tableau
                                {
                                    sb.Append("</tbody>");
                                    sb.Append("</table>");
                                    // saut de page + margin par rapport au début de la page pour éviter le chevauchement avec bandeau du haut de page
                                    sb.Append("<div style='page-break-before:always'>&nbsp;</div>");
                                    sb.Append("<div style='height:100px !important;min-height:100px !important;color:transparent;background-color:transparent;'>&nbsp;</div>");

                                    // démarrer un nouveau tableau dans une nouvelle page
                                    sb.Append("<table style='width:700px; font-size:12px;  text-align: center; repeat-header: yes; border:0px solid; border-collapse: collapse;'>");
                                    sb.Append(HeaderTable);
                                    sb.Append("<tbody>");
                                    NbPageVueGlobal++;
                                }
                                count++;
                            }
                        }
                        sb.Append("</tbody>");
                        sb.Append("</table></div>");
                        sb.Append("<span style='font-size:8px; font-family:Calibri;'>*SRRI : Synthetic Risk and Reward Indicator ou indicateur de risque et de performance </span><br/>" +
                            "<span style='font-size:8px; font-family:Calibri;'>**YTD: Year to date (performance depuis le 1er janvier)</ span>");
                        #endregion

                        #region Camember :
                        // Create first camembere :
                        Dictionary<string, double> dictcamembere1 = new Dictionary<string, double>();
                        Dictionary<string, double> dictTemporaire = new Dictionary<string, double>();
                        //Create seconde camembert : 
                        Dictionary<string, double> dictCamembertSupport1 = new Dictionary<string, double>();
                        Dictionary<string, double> dictCamembertSupportTemporaire = new Dictionary<string, double>();
                        double ValorisationTotal = 0;
                        if (releveSituation.Avoirs != null)
                        {
                            foreach (var r in releveSituation.Avoirs)
                            {
                                if (r.Positions != null)
                                {
                                    foreach (var rr in r.Positions)
                                    {
                                        if (dictcamembere1.ContainsKey(rr.ClasseActif))
                                        {
                                            dictcamembere1[rr.ClasseActif] += rr.Valorisation; // if element exist add the new value 
                                        }
                                        else
                                        {
                                            dictcamembere1.Add(rr.ClasseActif, rr.Valorisation);
                                        }

                                        // Seconde camembert :

                                        if (dictCamembertSupport1.ContainsKey(rr.ISIN))
                                        {
                                            dictCamembertSupport1[rr.ISIN] += rr.Valorisation; // if element exist add the new value 
                                        }
                                        else
                                        {
                                            dictCamembertSupport1.Add(rr.ISIN, rr.Valorisation);
                                        }

                                        ValorisationTotal += rr.Valorisation;
                                    }
                                }
                            }
                        }

                        //First camembert !!
                        foreach (var d in dictcamembere1)
                        {
                            dictTemporaire.Add(d.Key, (d.Value / ValorisationTotal) * 100);
                        }
                        dictcamembere1 = GetDictionnarySortedBy4(dictTemporaire);

                        // Seconde camembert : 
                        foreach (var d in dictCamembertSupport1)
                        {
                            dictCamembertSupportTemporaire.Add(d.Key, (d.Value / ValorisationTotal) * 100);
                        }
                        dictCamembertSupport1 = GetDictionnarySortedBy4(dictCamembertSupportTemporaire);

                        //generate camembert :
                        var FirstPathCamembere = GetGraphCamembert(dictcamembere1, Couleur_Base_Drawing);
                        var SecondePathCamembereMouvement = GetGraphCamembert(dictCamembertSupport1, Couleur_Base_Drawing);
                        ListPathToDelete.Add(FirstPathCamembere); // to delete in the end 
                        ListPathToDelete.Add(SecondePathCamembereMouvement); // to delete in the end 

                        // saut de page + margin par rapport au début de la page pour éviter le chevauchement avec bandeau du haut de page
                        if (releveSituation.Avoirs.Count > 6 && releveSituation.Avoirs.Count <= 12)
                        {
                            NbPageVueGlobal++;
                            sb.Append("<div style='page-break-before:always'>&nbsp;</div>");
                            sb.Append("<div style='height:60px !important;min-height:60px !important;color:transparent;background-color:transparent;'>&nbsp;</div>");
                        }
                        sb.Append("<div style='padding:15px;'><span style='font-size:18; font-weight: bold; font-family: Calibri;'>Répartition par classes d’actifs </span> <span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span>" +
                                  " <span style='margin-left:600px; font-size:18; font-weight: bold; font-family: Calibri;'>Répartition par supports </span></div>");
                        sb.Append($"<div> <img src='{FirstPathCamembere}'/> <img src='{SecondePathCamembereMouvement}'/> </div>");

                        #endregion

                        // New Page Mouvements:
                        #region Table Mouvements :

                        NbPageVueGlobal++;
                        sb.Append("<div style='page-break-before:always'>&nbsp;</div>");
                        sb.Append("<div style='height:60px !important;min-height:60px !important;color:transparent;background-color:transparent;'>&nbsp;</div>");
                        sb.Append("<span style='margin-left:600px; font-size:24; font-weight: bold; font-family: Calibri;'>Mouvements</span>");

                        var HeaderTableMouvements = "<thead> <tr style='height: 39px; color:white; font-family:Calibri;'> " +
                            "<th class='thBackground'> Compte/contrat </th> <th class='thBackground'> Type  </th> <th class='thBackground'> Montant </th> " +
                            "<th class='thBackground'> Quantité  </th> " +
                            "<th class='thBackground'>  Date d’effet </th> <th class='thBackground'> ISIN </th> <th class='thBackground'> Libellé </ th> " +
                            "</tr> </thead>";

                        sb.Append("<div><table style='width:700px; font-size:12px;  text-align: center; repeat-header: yes; border:0px solid; border-collapse: collapse;'>");
                        sb.Append(HeaderTableMouvements);

                        var countLine = 1;
                        sb.Append("<tbody>");
                        var NbMouvementGlobal = releveSituation.Avoirs.Sum(x => x.Mouvements.Count);
                        if (releveSituation.Avoirs != null)
                        {
                            foreach (var l in releveSituation.Avoirs)
                            {
                                if (l.Mouvements != null)
                                {
                                    foreach (var m in l.Mouvements)
                                    {
                                        if (countLine % 2 == 0)
                                        {
                                            sb.Append("<tr class='tdClass' style='height: 39px; background-color:#efeded; font-family:Calibri;'>");
                                        }
                                        else
                                        {
                                            sb.Append("<tr class='gris tdClass' style='height: 39px; background-color:#f9f9f9; font-family:Calibri;'>"); /*f1f1f1*/
                                        }

                                        sb.Append($"<td>{l.Nom}</td><td>{m.Type.ToString()}</td><td>{m.Montant.ToString("### ###")}€</td><td>{m.Quantite}</td><td>{m.DateEffet.ToString("dd/mm/yyy")}</td><td>{m.ISIN}</td><td>{m.Libelle}</td>");
                                        sb.Append("</tr>");
                                        if (countLine % 18 == 0 && countLine != NbMouvementGlobal) // la deuxieme condition c'est pour ne pas rajouter un saut de page apres le tableau
                                        {
                                            NbPageVueGlobal++;
                                            sb.Append("</tbody>");
                                            sb.Append("</table>");
                                            // saut de page + margin par rapport au début de la page pour éviter le chevauchement avec bandeau du haut de page
                                            sb.Append("<div style='page-break-before:always'>&nbsp;</div>");
                                            sb.Append("<div style='height:100px !important;min-height:100px !important;color:transparent;background-color:transparent;'>&nbsp;</div>");

                                            // démarrer un nouveau tableau dans une nouvelle page
                                            sb.Append("<table style='width:700px; font-size:12px;  text-align: center; repeat-header: yes; border:0px solid; border-collapse: collapse;'>");
                                            sb.Append(HeaderTableMouvements);
                                            sb.Append("<tbody>");
                                        }
                                        countLine++;
                                    }
                                }

                            }
                        }

                        sb.Append("</tbody>");
                        sb.Append("</table></div>");
                        #endregion

                        // Page After Mouvements:
                        NbPageVueGlobal++;
                        sb.Append("<div style='page-break-before:always'>&nbsp;</div>");
                        sb.Append("<div style='height:60px !important;min-height:60px !important;color:transparent;background-color:transparent;'>&nbsp;</div>");

                        var TableRisqueClient = GetTableOfRisque(releveSituation.RisqueClient, Couleur_Base_HEXA, "width:55px; ", "width:40px; ", "cellspacing='0'");
                        //Calcul le risque agrege des portfeuilles:
                        double RisqueAgrege = 0;
                        if (releveSituation.Avoirs != null)
                        {
                            foreach (var ag in releveSituation.Avoirs)
                            {
                                foreach (var p in ag.Positions)
                                {
                                    RisqueAgrege += (p.SRRI * p.Valorisation);
                                }
                            }
                        }

                        RisqueAgrege = RisqueAgrege / ValorisationTotal;

                        var TableRisqueAgregePortfeuille = GetTableOfRisque(RisqueAgrege, Couleur_Base_HEXA, "width:55px; ", "width:40px; ", "cellspacing='0'");

                        var DecalageWithoutMap = "";
                        var AddNewPageAfterVersementEtRetrait = "";
                        if (releveSituation.RepartitionGeographique != null && releveSituation.RepartitionGeographique.Count > 0)
                        {
                            NbPageVueGlobal++;
                            AddNewPageAfterVersementEtRetrait = "<div style='page-break-before:always'>&nbsp;</div><div style='height:100px !important;min-height:100px !important;color:transparent;background-color:transparent;'>&nbsp;</div>";
                            var PngMapPath = BuildWorldMap(releveSituation.RepartitionGeographique, Couleur_Base_HEXA);
                            ListPathToDelete.Add(PngMapPath); // to delete in the end 
                            sb.Append($"<div> <img src='{PngMapPath}' /> </div>");
                        }
                        else
                        {
                            DecalageWithoutMap = "<div style='height:50px;'> &nbsp;  </div>";  // faire de la place si y'a pas de map
                        }

                        sb.Append($"<div><div style='clear: both; float:left;'><span style='font-size:18; font-weight: bold; font-family: Calibri; '> Risque client </span><br/><br/>" + TableRisqueClient);
                        sb.Append($"<br/><span style='font-size:10; font-family: Calibri;'>Note pondérée de </span> <span style='font-size:10; font-family: Calibri; color:{Couleur_Base_HEXA};'> {Math.Round(releveSituation.RisqueClient, 2)}</span></div>");

                        sb.Append("<div style='clear: both; float:right;'> <span style='font-size:18; font-weight: bold; font-family: Calibri; '> Risque agrégé des portefeuilles </span><br/><br/> ");
                        sb.Append(TableRisqueAgregePortfeuille);
                        sb.Append($"<br/><span style='font-size:10; font-family: Calibri;'>Note pondérée de </span> <span style='font-size:10; font-family: Calibri; color:{Couleur_Base_HEXA};'> {Math.Round(RisqueAgrege, 2)}</span>");
                        sb.Append("</div></div>");

                        sb.AppendLine(DecalageWithoutMap); // faire de la place si y'a pas de map
                        sb.Append($"<div style='margin-left:550px; padding-bottom:15px;'> <span style='font-size:18; font-weight: bold; font-family: Calibri; '>Versements et retraits </span>" +
                            $"<span style='font-size:12; font-family: Calibri;'> sur la période du  {releveSituation.PeriodeDebut.ToString("dd/MM/yyy")} au {releveSituation.PeriodeFin.ToString("dd/MM/yyy")}</span></div>");

                        List<Tuple<string, double, double>> ListGraphBar = GetListVersementRetrait(releveSituation);
                        var GraphVersementRetrait = GetGraphBar(ListGraphBar, Couleur_Base_Drawing);
                        ListPathToDelete.Add(GraphVersementRetrait); // to delete in the end 

                        List<Tuple<double[], string>> LineGraph = GetListEvolutionPatrimoine(releveSituation);
                        var GrapheEvolutionPatrimoine = GetGraphLine(LineGraph, releveSituation, Couleur_Base_Drawing, 200, true);
                        ListPathToDelete.Add(GrapheEvolutionPatrimoine); // to delete in the end 

                        sb.Append($"<img src='{GraphVersementRetrait}'/>");
                        sb.Append(AddNewPageAfterVersementEtRetrait); // add new page if we have map graph !
                        sb.Append($"<div style='margin-left:550px; padding-bottom:15px;'> <span style='font-size:18; font-weight: bold; font-family: Calibri; '>Evolution du patrimoine </span>" +
                            $"<span style='font-size:12; font-family: Calibri;'> - Valorisation du {releveSituation.PeriodeDebut.ToString("dd/MM/yyy")} au {releveSituation.PeriodeFin.ToString("dd/MM/yyy")}</span></div>");
                        sb.Append($"<img src='{GrapheEvolutionPatrimoine}'/>");

                        //6eme Page : SITUATION DU CONTRAT NOVAVIE ACTIF PLUS
                        var NbPageFirstContrat = NbPageVueGlobal + 1;
                        sb.Append("<div style='page-break-before:always'>&nbsp;</div>");
                        sb.Append("<div style='height:60px !important;min-height:60px !important;color:transparent;background-color:transparent;'>&nbsp;</div>");
                        //saut de page


                        //Partie par contrats : 
                        if (releveSituation.Avoirs.Count() > 0 && releveSituation.Avoirs != null)
                        {
                            var ListPathClientsToDelete = new List<string>();
                            var Clients = GetSituationParAvoir(releveSituation, Couleur_Base_HEXA, Couleur_Base_Drawing, out ListPathClientsToDelete, NbPageFirstContrat, out ListSommaire);
                            ListPathToDelete.AddRange(ListPathClientsToDelete);
                            sb.Append(Clients);
                        }

                        //----------------- frais, couts et charges :
                        var HeaderTableFraiscoutsCharges = "<thead> <tr style='height: 39px; color:white; font-family:Calibri; font-size:11;  '> " +
                                                   "<th class='thBackground' style='padding-left:10px;'> Type  </th> " +
                                                   "<th class='thBackground' style='text-align:right; padding-right:20px; '> Montant</th> " +
                                                   "<th class='thBackground' style='padding-right:20px;'> Pourcentage </th> " +
                                                   "<th class='thBackground' style='padding-left:10px;'> Date</th> " +
                                                   "</tr> </thead>";
                        sb.Append("<br/><br/><span style='margin-left:600px; font-size:24; font-weight: bold; font-family: Calibri;'>Frais, coûts et charges </span>");

                        var firstTable = releveSituation.Frais.Take(11);
                        sb.AppendLine(BuildFraisTable(firstTable, HeaderTableFraiscoutsCharges, Couleur_Base_HEXA, true));
                        var NbPageContainerFrais = 0; // for sommaire 
                        if (releveSituation.Frais.Count > 11)
                        {
                            var FraisCollections = Helpers.SplitFile<Frais>(releveSituation.Frais.Skip(11), 20);


                            var countOperationPositionTable = 1;
                            if (releveSituation.Frais.Count > 6)
                            {
                                // saut de page + margin par rapport au début de la page pour éviter le chevauchement avec bandeau du haut de page
                                sb.AppendLine("<div style='page-break-before:always'>&nbsp;</div>");
                                sb.AppendLine("<div style='height:100px !important;min-height:100px !important;color:transparent;background-color:transparent;'>&nbsp;</div>");
                            }

                            foreach (var c in FraisCollections)
                            {
                                NbPageContainerFrais++; // sommaire
                                if (countOperationPositionTable == FraisCollections.Count())
                                {
                                    sb.AppendLine(BuildFraisTable(c, HeaderTableFraiscoutsCharges, Couleur_Base_HEXA, true));
                                }
                                else
                                {
                                    sb.AppendLine(BuildFraisTable(c, HeaderTableFraiscoutsCharges, Couleur_Base_HEXA));
                                }
                                countOperationPositionTable++;
                            }
                        }

                        // saut de page + margin par rapport au début de la page pour éviter le chevauchement avec bandeau du haut de page
                        sb.AppendLine("<div style='page-break-before:always'>&nbsp;</div>");
                        sb.AppendLine("<div style='height:100px !important;min-height:100px !important;color:transparent;background-color:transparent;'>&nbsp;</div>");
                        var NbPageDocumentInformation = 1 + NbPageContainerFrais + (ListSommaire.Count > 0 ? ListSommaire.Last().Item2 : 0);
                        ListSommaire.Add(Tuple.Create("Document d'information non contractuel", NbPageDocumentInformation));

                        //----------------- DOCUMENT D'INFORMATION NON CONTRACTUEL
                        sb.Append("<span style='margin-left:600px; font-size:24; font-weight: bold; font-family: Calibri;'>DOCUMENT D'INFORMATION NON CONTRACTUEL </span>");
                        sb.AppendLine("<div style='height:700px;  background-color:#F7F7F7'>" +
                            "<p style='padding-top:60px; padding-left:60px; padding-right:60px; font-size:17; font-weight: bold; font-family: Calibri;'>" +
                            "Ce document est destiné exclusivement à des clients non professionnels au sens de la Directive MIF. Il ne peut être utilisé dans un but autre que celui pour lequel il a été conçu et ne peut pas être reproduit, diﬀusé ou communiqué à des tiers en tout ou partie sans l’autorisation préalable et écrite de l'émetteur du document.<br/> Aucune information contenue dans ce document ne saurait être interprétée comme possédant une quelconque valeur contractuelle. Ce document est produit à titre purement indicatif. Il constitue une présentation conçue et réalisée par l'émetteur du document à partir de sources qu’il estime ﬁables. L'émetteur du document se réserve la possibilité de modiﬁer les informations présentées dans ce document à tout moment et sans préavis qui ne constitue en aucun cas un engagement de la part de celui-ci. L'émetteur du document ne saurait être tenue responsable de toute décision prise ou non sur la base d’une information contenue dans ce document, ni de l’utilisation qui pourrait en être faite par un tiers.<br/> <br/> Les illustrations graphiques présentées ne constituent pas une prévision de la performance future de vos investissements.Elles ont seulement pour but d’illustrer les mécanismes de votre investissement sur la durée de placement.L’évolution de la valeur de votre investissement pourra s’écarter de ce qui est aﬃché, à la hausse comme à la baisse.Les gains et les pertes peuvent dépasser les montants aﬃchés, respectivement, dans les scénarios les plus favorables et les plus défavorables.<br/>En utilisant ce document, vous reconnaissez avoir pris connaissance de cet avertissement, l’avoir compris et en accepter le contenu.Les simulations sont réalisées à titre d'exemple sur la base d'hypothèses de gestion ﬁnancière en date du 05 / 04 / 2018, elles ne constituent pas un engagement contractuel et ne sauraient engager la responsabilité de l'émetteur du document. <br/> Nous tenons à votre disposition les hypothèses utilisées. Les performances passées ne préjugent pas des performances futures." +
                            " </p>" +
                            "</div>");

                        sb.Append("</body></html>");

                        var example_css = @".hei{height:42%}.thBackground{background-color: [background_color_property];}.gris{background-color:#f1f1f1; }.tdRisqueBorderWhite{ border-right:2px solid; border-right-color:white;}.tdClass{color: black;}th{min-width:30px;}.SpanTitle{color: [background_color_property]; font-size:24;  font-weight: bold; font-family: Calibri;}.TDClients{border-width:absolute; border-right:1px solid; border-right-color:appworkspace; border-right-width:20px; width:160px; height:78px; text-align:center; color:white; background-color:[background_color_property]; font-size:18; font-weight:bold; font-family:Calibri;}.TDClientsGrey{width:179;background-color:#f9f9f9;color:black;height:48px; text-align:center;font-size:14; font-family:Calibri;}";
                        var Css_value = example_css.Replace("[background_color_property]", Couleur_Base_HEXA);

                        using (var msCss = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(Css_value)))
                        {
                            using (var msHtml = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(sb.ToString())))
                            {
                                //Parse the HTML
                                iTextSharp.tool.xml.XMLWorkerHelper.GetInstance().ParseXHtml(writer, doc1, msHtml, msCss);
                            }
                        }
                        doc1.Close();
                    }

                    #region Traitement Header + Footer + Rename File + Delete temporary files : 

                    foreach (var item in ListPathToDelete)
                    {
                        File.Delete(item);
                    }

                    var PathPageDeGarde = $"PageDeGarde{DateTime.Now:ddMMyyyy_HH - mm - ss}";
                    File.Copy(Path.Combine(Directory, outFile), Path.Combine(Directory, PathPageDeGarde));
                    File.Delete(Path.Combine(Directory, outFile));//Delete Page de garde
                    File.Copy(Path.Combine(Directory, TemporaryDoc), Path.Combine(Directory, outFile)); // generate final document with outPutPath of param
                    File.Delete(Path.Combine(Directory, TemporaryDoc)); //Delete Page document of all page except page de garde without footer and header

                    //Header and footer are inside InsertSommaire
                    InsertSommaire(Path.Combine(Directory, outFile), FontSommaire, ListSommaire, releveSituation, PathPageDeGarde);

                    #endregion


                }
                catch (Exception e)
                {
                    throw;
                }
            }
            return true;
        }

        private static string GetSituationParAvoir(ReleveSituation releveSituation, string Couleur_Base_HEXA, System.Drawing.Color Couleur_Base_Drawing, out List<string> ListPathToDeleteEnd
                                                , int NbPageStart, out List<Tuple<string, int>> ListSommaire)
        {
            try
            {
                var ListSommaireTemp = new List<Tuple<string, int>>();
                var NbPageStartParContrat = NbPageStart;
                StringBuilder sb = new StringBuilder();
                var ListPathToDelete = new List<string>(); // List path to delete in the end
                var TitulaireDistinct = releveSituation.Avoirs.GroupBy(l => l.Titulaire).Select(g => g.First()).ToList();
                var count = 1;
                foreach (var L in releveSituation.Avoirs)
                {
                    ListSommaireTemp.Add(Tuple.Create(L.Nom, NbPageStartParContrat)); // calcul sommaire 
                    sb.AppendLine($"<br/><span class='SpanTitle'>SITUATION DU CONTRAT {L.Nom} </ span>");

                    #region Partie Titulaire :

                    sb.AppendLine($"<br/><br/><span style='font-size:12; font-family: Calibri;'> Titulaires :  </span>");
                    var CountTitulaire = 1;
                    foreach (var a in TitulaireDistinct)
                    {
                        if (CountTitulaire != TitulaireDistinct.Count)
                            sb.AppendLine($"<span style='font-size:12; font-family: Calibri; color:{Couleur_Base_HEXA};'> {a.Titulaire} - </span>");
                        else
                            sb.AppendLine($"<span style='font-size:12; font-family: Calibri; color:{Couleur_Base_HEXA};'> {a.Titulaire} </span>");

                        CountTitulaire++;
                    }
                    #endregion

                    #region Panel
                    //Panel :
                    var PathImgPerformanceOrigine = L.PerfOrigine > 0 ? PathUpImage : L.PerfOrigine < 0 ? PathDownImage : PathEqualImage;
                    var PathImgPerformance1Janvier = L.PerfYTD > 0 ? PathUpImage : L.PerfYTD < 0 ? PathDownImage : PathEqualImage;
                    sb.AppendLine($"<table> <tr> " +
                              $" <td rowspan='2' class='TDClients' style='font-weight:bold;'> Etablissement <br/>{L.Etablissement}</td>" +
                              $" <td rowspan='2' class='TDClients' style='font-weight:bold;'> Valorisation  <br/>{L.Valorisation.ToString("### ###")}€</td> " +
                              $" <td rowspan='2' class='TDClients' style='font-weight:bold;'> Performance depuis l'origine <br/>  <img src='{PathImgPerformanceOrigine}'  width='17' height='17' />  &nbsp;&nbsp;{L.PerfOrigine.ToString("##.#")}% </td>" +
                              $" <td rowspan='2' class='TDClients' style='font-weight:bold;'> Performance depuis le  1 Jan  <br/> <img src='{PathImgPerformance1Janvier}' width='17' height='17' /> &nbsp;&nbsp;{L.PerfYTD.ToString("##.#")}% </td> " +
                              $"</tr></table>");
                    var sommeRetrait = L.SommeRetraits == 0 ? "0" : L.SommeRetraits.ToString("### ###");
                    var sommeVersement = L.SommeVersements == 0 ? "0" : L.SommeVersements.ToString("### ###");
                    sb.AppendLine("<div style='height:10px;'> &nbsp;&nbsp; </div>");
                    sb.AppendLine($" <table style='border-collapse: collapse;'> <tr>" +
                              $" <td class='TDClientsGrey'> Numéro <br/> <span style='color:{Couleur_Base_HEXA};'> {L.Numero} </span></td>" +
                              $" <td class='TDClientsGrey'> Type   <br/> <span style='color:{Couleur_Base_HEXA};'> {L.Type} </span></td> " +
                              $" <td class='TDClientsGrey'> Somme des versements <br/> <span style='color:{Couleur_Base_HEXA};'> {sommeVersement}€ </span> </td>" +
                              $" <td class='TDClientsGrey'> Somme des retraits  <br/>  <span style='color:{Couleur_Base_HEXA};'> {sommeRetrait}€ </span></td> " +
                              $"</tr> </table>");

                    var Pmv = L.PMV == 0 ? "0" : L.PMV.ToString("### ###");
                    sb.AppendLine("<div style='height:10px;'> &nbsp;&nbsp; </div>");
                    sb.AppendLine($" <table style='border-collapse: collapse;'> <tr>" +
                              $" <td class='TDClientsGrey'> Date d’ouverture <br/> <span style='color:{Couleur_Base_HEXA};'> {L.DateOuverture.ToString("dd/mm/yyy")} </span></td>" +
                              $" <td class='TDClientsGrey'> Moyenne SRRI*    <br/> <span style='color:{Couleur_Base_HEXA};'> {Math.Round(L.SRRIMoyen, 2)} </span></td> " +
                              $" <td class='TDClientsGrey'> +/- Value <br/> <span style='color:{Couleur_Base_HEXA};'> {Pmv}€ </span> </td>" +
                              $" <td class='TDClientsGrey'> Performance entre le {releveSituation.PeriodeDebut.ToString("dd MMM yyyy", CultureInfo.CreateSpecificCulture("fr-FR"))}" +
                              $" et le {releveSituation.PeriodeFin.ToString("dd MMM yyyy", CultureInfo.CreateSpecificCulture("fr-FR"))}  " +
                              $" <br/>  <span style='color:{Couleur_Base_HEXA};'> {L.PerfPeriode} </span></td> " +
                              $"</tr> </table>");
                    #endregion

                    //Echelle de risque client:
                    sb.AppendLine($" <br/> <span style='font-size:24; font-weight:bold; font-family:Calibri;'> Echelle de risque client </span>");
                    sb.AppendLine(GetTableOfRisque(L.SRRIMoyen, Couleur_Base_HEXA, "width:140px; ", "width:95px; ", "cellspacing='0'"));
                    sb.AppendLine($"<span style='font-size:10px; font-family:Calibri;'> Note pondérée de </span> <span  style='color:{Couleur_Base_HEXA}; font-size:10px; font-family:Calibri;'>{Math.Round(L.SRRIMoyen, 2)} </span> <br/>");

                    sb.AppendLine("<br/><span style='font-size:24; font-weight:bold; font-family:Calibri;'> Positions  </span>");
                    //--- Table Positions :
                    var HeaderTablePositions = "<thead> <tr style='height: 39px; color:white; font-family:Calibri; font-size:11;'> " +
                   "<th class='thBackground'> ISIN  </th> " +
                   "<th class='thBackground'> Libellé  </th> " +
                   "<th class='thBackground'> Quantité </th> " +
                   "<th class='thBackground'> Cours  </th> " +
                   "<th class='thBackground'> Valorisation  </th> " +
                   "<th class='thBackground'> PAM*  </th> " +
                   "<th class='thBackground'> Poids  </th> " +
                   "<th class='thBackground'> +/- value en € </ th> " +
                   "<th class='thBackground'> +/- value en % </ th> " +
                   "<th class='thBackground'> Date du cours  </th> " +
                   "<th class='thBackground'>  SRRI** </th> " +
                   "</tr> </thead>";

                    var firstTable = L.Positions.Take(6);
                    sb.AppendLine(BuildPositionsTable(firstTable, HeaderTablePositions, Couleur_Base_HEXA, true));

                    if (L.Positions.Count > 6)
                    {
                        var positionsCollections = Helpers.SplitFile<Position>(L.Positions.Skip(6), 20);
                        var countOperationPositionTable = 1;

                        // saut de page + margin par rapport au début de la page pour éviter le chevauchement avec bandeau du haut de page
                        sb.AppendLine("<div style='page-break-before:always'>&nbsp;</div>");
                        sb.AppendLine("<div style='height:100px !important;min-height:100px !important;color:transparent;background-color:transparent;'>&nbsp;</div>");

                        foreach (var c in positionsCollections)
                        {
                            NbPageStartParContrat++; // calcul nb page par contrat
                            if (countOperationPositionTable == positionsCollections.Count())
                            {
                                sb.AppendLine(BuildPositionsTable(c, HeaderTablePositions, Couleur_Base_HEXA, true));
                            }
                            else
                            {
                                sb.AppendLine(BuildPositionsTable(c, HeaderTablePositions, Couleur_Base_HEXA));
                            }
                            countOperationPositionTable++;
                        }
                    }

                    sb.AppendLine("<div style='page-break-before:always'>&nbsp;</div>");
                    sb.AppendLine("<div style='height:100px !important;min-height:100px !important;color:transparent;background-color:transparent;'>&nbsp;</div>");
                    NbPageStartParContrat++; // calcul nb page par contrat

                    //-------------- Mouvements par contrats :
                    if (L.Mouvements.Count() > 0 && L.Mouvements != null)
                    {
                        var NbPageMouvement = 0;
                        sb.AppendLine(BuildMouvemantTable(L.Mouvements, ref NbPageMouvement));
                        NbPageStartParContrat += NbPageMouvement;
                    }

                    //-------------- Performance : 
                    sb.Append("<span style='margin-left:600px; font-size:24; font-weight: bold; font-family: Calibri;'>Performances</ span>");
                    var PathImageDepuisLePerformance = L.PerfYTD > 0 ? PathUpImage : L.PerfYTD < 0 ? PathDownImage : PathEqualImage;
                    var PathImagePerformanceMoins1 = L.PerfNMoins1 > 0 ? PathUpImage : L.PerfNMoins1 < 0 ? PathDownImage : PathEqualImage;
                    var PathImagePerformanceMoins2 = L.PerfNMoins2 > 0 ? PathUpImage : L.PerfNMoins2 < 0 ? PathDownImage : PathEqualImage;
                    var PathImagePerformanceMoins3 = L.PerfNMoins3 > 0 ? PathUpImage : L.PerfNMoins3 < 0 ? PathDownImage : PathEqualImage;
                    sb.AppendLine($"<table> <tr> " +
                             $" <td rowspan='2' class='TDClients'><span style='font-size:10;'> Depuis le <br/> 01/01/{L.DateOuverture.Year} </span><br/> <img src='{PathImageDepuisLePerformance}'  width='17' height='17'/> &nbsp;&nbsp; {L.PerfYTD} %</td>" +
                             $" <td rowspan='2' class='TDClients'> N-1 <br/> <img src='{PathImagePerformanceMoins1}'  width='17' height='17' />  &nbsp;&nbsp; {L.PerfNMoins1} %</td> " +
                             $" <td rowspan='2' class='TDClients'> N-2 <br/> <img src='{PathImagePerformanceMoins2}'  width='17' height='17' />  &nbsp;&nbsp; {L.PerfOrigine} % </td>" +
                             $" <td rowspan='2' class='TDClients'> N-3 <br/> <img src='{PathImagePerformanceMoins3}'  width='17' height='17' />  &nbsp;&nbsp; {L.PerfYTD} % </td> " +
                             $"</tr></table>");

                    //--------------- Camembere : 
                    if (L.Positions.Count() > 0 && L.Positions != null)
                    {
                        var ListTempPathCamembere = new List<string>();
                        sb.AppendLine(Get2Camembere(L.Positions, Couleur_Base_Drawing, out ListTempPathCamembere));
                        ListPathToDelete.AddRange(ListTempPathCamembere);
                    }

                    //--------------- Courbe :
                    #region LineGraphe Par Contrats :

                    var NbElement = L.HistoriqueValorisations.Where(h => h.Key.Date >= releveSituation.PeriodeDebut.Date && h.Key.Date <= releveSituation.PeriodeFin.Date);
                    if (NbElement.Count() > 0)
                    {
                        var Points = new double[NbElement.Count()];
                        var P = 0;
                        foreach (var d in L.HistoriqueValorisations.OrderBy(x => x.Key))
                        {
                            if (d.Key.Date >= releveSituation.PeriodeDebut.Date && d.Key.Date <= releveSituation.PeriodeFin.Date) // La periode doit etre dans la periode global 
                            {
                                Points[P] = d.Value;
                            }
                            P++;
                        }
                        List<Tuple<double[], string>> LineGraph = new List<Tuple<double[], string>>() { Tuple.Create(Points, L.Nom) };
                        var GrapheEvolutionPatrimoine = GetGraphLine(LineGraph, releveSituation, Couleur_Base_Drawing, 250);
                        ListPathToDelete.Add(GrapheEvolutionPatrimoine); // delete in the end of generatiion of document
                        sb.Append($"<br/><div style='margin-left:550px; padding-bottom:15px;'> <span style='font-size:24; font-weight: bold; font-family: Calibri; '>Evolution du patrimoine </span>" +
                      $"<span style='font-size:12; font-family: Calibri;'> sur la période du {releveSituation.PeriodeDebut.ToString("dd/MM/yyy")} au {releveSituation.PeriodeFin.ToString("dd/MM/yyy")}</span></div>");
                        sb.Append($"<img src='{GrapheEvolutionPatrimoine}'/>");
                    }

                    #endregion

                    //------ Add New page : 
                    sb.AppendLine("<div style='page-break-before:always'>&nbsp;</div>");
                    sb.AppendLine("<div style='height:100px !important;min-height:100px !important;color:transparent;background-color:transparent;'>&nbsp;</div>");
                    NbPageStartParContrat++; // calcul nb page par contrat

                    //---------------- Versements et retraits Graphe :
                    sb.Append($"<div style='margin-left:550px; padding-bottom:15px;'> <span style='font-size:24; font-weight: bold; font-family: Calibri; '>Versements et retraits </span>" +
                               $"<span style='font-size:12; font-family: Calibri;'> sur la période du  {releveSituation.PeriodeDebut.ToString("dd/MM/yyy")} au {releveSituation.PeriodeFin.ToString("dd/MM/yyy")}</span></div>");

                    List<Tuple<string, double, double>> ListGraphBar = GetListVersementRetrait(releveSituation, L);
                    var GraphVersementRetrait = GetGraphBar(ListGraphBar, Couleur_Base_Drawing, 220);
                    ListPathToDelete.Add(GraphVersementRetrait); // delete in the end 
                    sb.Append($"<img src='{GraphVersementRetrait}'/>");


                    //--------------- Separation par contrats :
                    if (count != releveSituation.Avoirs.Count)
                    {
                        NbPageStartParContrat++; // calcul nb page par contrat
                        sb.AppendLine("<div style='page-break-before:always'>&nbsp;</div>");
                        sb.AppendLine("<div style='height:100px !important;min-height:100px !important;color:transparent;background-color:transparent;'>&nbsp;</div>");
                    }
                    count++;
                }
                ListSommaireTemp.Add(Tuple.Create("Synthèse des frais, coûts et charges ", NbPageStartParContrat)); // calcul Nb page start Frais !
                ListPathToDeleteEnd = ListPathToDelete;
                ListSommaire = ListSommaireTemp;
                return sb.ToString();
            }
            catch (Exception)
            {

                throw;
            }

        }

        private static void GetDictionnary(List<Position> Positions, ref Dictionary<string, double> dictcamembere1, ref Dictionary<string, double> dictCamembertSupport1, ref double ValorisationTotal)
        {
            try
            {
                foreach (var rr in Positions)
                {
                    if (dictcamembere1.ContainsKey(rr.ClasseActif))
                    {
                        dictcamembere1[rr.ClasseActif] += rr.Valorisation; // if element exist add the new value 
                    }
                    else
                    {
                        dictcamembere1.Add(rr.ClasseActif, rr.Valorisation);
                    }

                    // Seconde camembert :

                    if (dictCamembertSupport1.ContainsKey(rr.ISIN))
                    {
                        dictCamembertSupport1[rr.ISIN] += rr.Valorisation; // if element exist add the new value 
                    }
                    else
                    {
                        dictCamembertSupport1.Add(rr.ISIN, rr.Valorisation);
                    }

                    ValorisationTotal += rr.Valorisation;
                }
            }
            catch (Exception)
            {

                throw;
            }

        }
        private static string Get2Camembere(List<Position> Positions, System.Drawing.Color Couleur_Base_Drawing, out List<string> ListPathToDelete)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                var listTempToDelete = new List<string>();
                // Create first camembere :
                Dictionary<string, double> dictcamembere1 = new Dictionary<string, double>();
                Dictionary<string, double> dictTemporaire = new Dictionary<string, double>();
                //Create seconde camembert : 
                Dictionary<string, double> dictCamembertSupport1 = new Dictionary<string, double>();
                Dictionary<string, double> dictCamembertSupportTemporaire = new Dictionary<string, double>();
                double ValorisationTotal = 0;

                //cas de chaque contrats tout seul :
                GetDictionnary(Positions, ref dictcamembere1, ref dictCamembertSupport1, ref ValorisationTotal);


                //First camembert !!
                foreach (var d in dictcamembere1)
                {
                    dictTemporaire.Add(d.Key, (d.Value / ValorisationTotal) * 100);
                }
                dictcamembere1 = GetDictionnarySortedBy4(dictTemporaire);

                // Seconde camembert : 
                foreach (var d in dictCamembertSupport1)
                {
                    dictCamembertSupportTemporaire.Add(d.Key, (d.Value / ValorisationTotal) * 100);
                }
                dictCamembertSupport1 = GetDictionnarySortedBy4(dictCamembertSupportTemporaire);


                //generate camembert :
                var FirstPathCamembere = GetGraphCamembert(dictcamembere1, Couleur_Base_Drawing);
                var SecondePathCamembereMouvement = GetGraphCamembert(dictCamembertSupport1, Couleur_Base_Drawing);

                listTempToDelete.Add(FirstPathCamembere);
                listTempToDelete.Add(SecondePathCamembereMouvement);
                ListPathToDelete = listTempToDelete;

                sb.Append("<div style='padding:50px;'><span style='font-size:20; font-weight: bold; font-family: Calibri;'>Répartition par classes d’actifs </span> <span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span>" +
                          " <span style='margin-left:600px; font-size:20; font-weight: bold; font-family: Calibri;'>Répartition par supports </span></div>");
                sb.Append($"<div> <img src='{FirstPathCamembere}' style='width:340;'/> <span style='width:4px;'> &nbsp;&nbsp; </span> <img src='{SecondePathCamembereMouvement}'  style='width:340;'/> </div>");
                return sb.ToString();
            }
            catch (Exception)
            {

                throw;
            }

        }
        private static string BuildMouvemantTable(List<Mouvement> mouvements, ref int nbPageMouvements)
        {
            try
            {
                StringBuilder sb = new StringBuilder();

                sb.Append("<span style='margin-left:600px; font-size:24; font-weight: bold; font-family: Calibri;'>Mouvements</span>");

                var HeaderTableMouvements = "<thead> <tr style='height: 39px; color:white; font-family:Calibri; font-size:11;'> " +
                    " <th class='thBackground'> Type  </th> " +
                    "<th class='thBackground'> Montant </th> " +
                    "<th class='thBackground'> Quantité  </th> " +
                    "<th class='thBackground'>  Date d’effet </th> " +
                    "<th class='thBackground'> ISIN </th> " +
                    "<th class='thBackground'> Libellé </ th> " +
                    "</tr> </thead>";

                sb.Append("<div><table style='width:720px; font-size:12px;  text-align: center; repeat-header: yes; border:0px solid; border-collapse: collapse;'>");
                sb.Append(HeaderTableMouvements);

                //var NbPagetemp = 0;
                var countLine = 1;
                sb.Append("<tbody>");
                foreach (var m in mouvements)
                {
                    if (countLine % 2 == 0)
                    {
                        sb.Append("<tr class='tdClass' style='height: 39px; background-color:#F0F1F1; font-family:Calibri;'>");
                    }
                    else
                    {
                        sb.Append("<tr class='gris tdClass' style='height: 39px; background-color:#F7F7F7; font-family:Calibri;'>");
                    }

                    var Montant = m.Montant == 0 ? "0" : m.Montant.ToString("### ###");
                    sb.Append($"<td>{m.Type}</td><td>{Montant} € </td><td>{m.Quantite}</td><td>{m.DateEffet.ToString("dd/mm/yyy")}</td><td>{m.ISIN}</td><td>{m.Libelle}</td>");
                    sb.Append("</tr>");
                    if (countLine % 18 == 0 && countLine != mouvements.Count) // la deuxieme condition c'est pour ne pas rajouter un saut de page apres le tableau
                    {
                        //NbPagetemp++;
                        nbPageMouvements++;
                        sb.Append("</tbody>");
                        sb.Append("</table>");
                        // saut de page + margin par rapport au début de la page pour éviter le chevauchement avec bandeau du haut de page
                        sb.Append("<div style='page-break-before:always'>&nbsp;</div>");
                        sb.Append("<div style='height:100px !important;min-height:100px !important;color:transparent;background-color:transparent;'>&nbsp;</div>");

                        // démarrer un nouveau tableau dans une nouvelle page
                        sb.Append("<table style='width:720px; font-size:12px;  text-align: center; repeat-header: yes; border:0px solid; border-collapse: collapse;'>");
                        sb.Append(HeaderTableMouvements);
                        sb.Append("<tbody>");
                    }
                    countLine++;
                }

                sb.Append("</tbody>");
                sb.Append("</table></div>");
                sb.AppendLine("<div style='page-break-before:always'>&nbsp;</div>");
                sb.AppendLine("<div style='height:100px !important;min-height:100px !important;color:transparent;background-color:transparent;'>&nbsp;</div>");
                //NbPagetemp++;
                nbPageMouvements++;
                //nbPageMouvements = NbPagetemp;
                return sb.ToString();
            }
            catch (Exception)
            {
                throw;
            }

        }

        public static string BuildPositionsTable(IEnumerable<Position> positions, string tableHeader, string Couleur_Base_HEXA, bool LastOperation = false)
        {
            try
            {
                var sb = new StringBuilder();
                var countLineTablePosition = 1;
                sb.AppendLine("<table style='width:720px; font-size:12px;  text-align: center; repeat-header: yes; border:0px solid; border-collapse: collapse;'>");
                sb.AppendLine(tableHeader);
                sb.AppendLine("<tbody>");
                foreach (var p in positions)
                {
                    if (countLineTablePosition % 2 == 0)
                    {
                        sb.AppendLine("<tr class='tdClass' style='height: 39px; background-color:#efeded; font-family:Calibri;'>");
                    }
                    else
                    {
                        sb.AppendLine("<tr class='gris tdClass' style='height: 39px; background-color:#f9f9f9; font-family:Calibri;'>");
                    }
                    var couleurValuePourcentage = p.PMVPourcent < 0 ? "red" : Couleur_Base_HEXA;
                    var couleurValueEuro = p.PMVEuros < 0 ? "red" : Couleur_Base_HEXA;

                    var PmvEuro = p.PMVEuros == 0 ? "0" : p.PMVEuros.ToString("### ###");
                    sb.AppendLine($"<td>{p.ISIN}</td><td>{p.Libelle}</td><td>{p.Quantite.ToString("###.#")}</td><td>{p.Cours.ToString("###.#")}</td><td>{p.Valorisation.ToString("### ###.#")}</td>" +
                        $"<td>{p.PAM.ToString("###.#")}</td><td>{p.Poids.ToString("###.#")}</td>" +
                        $"<td><span style='color:{couleurValueEuro};'> {PmvEuro}€ </span></td>" +
                        $"<td><span style='color:{couleurValuePourcentage};'> {p.PMVPourcent.ToString("##.#")} % </span></td>" +
                        $"<td> {p.DateCours.ToString("dd/mm/yyy")}</td><td> {p.SRRI}</td>");
                    sb.AppendLine("</tr>");
                    countLineTablePosition++;
                }
                sb.AppendLine("</tbody>");
                sb.AppendLine("</table>");

                if (!LastOperation)
                {
                    // saut de page + margin par rapport au début de la page pour éviter le chevauchement avec bandeau du haut de page
                    sb.AppendLine("<div style='page-break-before:always'>&nbsp;</div>");
                    sb.AppendLine("<div style='height:100px !important;min-height:100px !important;color:transparent;background-color:transparent;'>&nbsp;</div>");
                }
                else
                {
                    sb.AppendLine("<span style='font-size:8px; font-family:Calibri;'> *PAM : Prix d’achat moyen</span>");
                    sb.AppendLine("<br/><span style='font-size:8px; font-family:Calibri;'>**SRRI : Synthetic Risk and Reward Indicator ou indicateur de risque et de performance </span>");
                }
                return sb.ToString();
            }
            catch (Exception e)
            {
                throw;
            }
        }


        public static string BuildFraisTable(IEnumerable<Frais> frais, string tableHeader, string Couleur_Base_HEXA, bool LastOperation = false)
        {
            try
            {
                var sb = new StringBuilder();
                var countLineTablePosition = 1;
                sb.AppendLine("<table style='width:720px; font-size:12px;  text-align: center; repeat-header: yes; border:0px solid; border-collapse: collapse;'>");
                sb.AppendLine(tableHeader);
                sb.AppendLine("<tbody>");
                foreach (var p in frais)
                {
                    if (countLineTablePosition % 2 == 0)
                    {
                        sb.AppendLine("<tr class='tdClass' style='height: 39px; background-color:#efeded; font-family:Calibri;'>");
                    }
                    else
                    {
                        sb.AppendLine("<tr class='gris tdClass' style='height: 39px; background-color:#f9f9f9; font-family:Calibri;'>");
                    }
                    var Montant = p.Montant == 0 ? "0" : p.Montant.ToString("### ###");
                    var Pourcentage = p.Pourcentage == 0 ? "0" : p.Pourcentage.ToString("##.#");
                    sb.AppendLine($"<td style='padding-left:10px;'>{p.Type}</td>" +
                     $"<td style='padding-right:20px; text-align:right;'>{Montant}€</td>" +
                     $"<td style='padding-left:10px;'>{Pourcentage} % </td>" +
                     $"<td>{p.Date.ToString("dd/mm/yyy")}</td>");

                    sb.AppendLine("</tr>");
                    countLineTablePosition++;
                }
                sb.AppendLine("</tbody>");
                sb.AppendLine("</table>");

                if (!LastOperation)
                {
                    // saut de page + margin par rapport au début de la page pour éviter le chevauchement avec bandeau du haut de page
                    sb.AppendLine("<div style='page-break-before:always'>&nbsp;</div>");
                    sb.AppendLine("<div style='height:100px !important;min-height:100px !important;color:transparent;background-color:transparent;'>&nbsp;</div>");
                }
                return sb.ToString();
            }
            catch (Exception e)
            {
                throw;
            }

        }

        private static List<Tuple<string, double, double>> GetListVersementRetrait(ReleveSituation releveSituation, Avoir avoir = null) // deuxieme param utilisé pour les graphe par contrats pas besoin de boucles sur la liste des avoirs 
        {
            try
            {
                List<Tuple<string, double, double>> ListGraphBar = new List<Tuple<string, double, double>>();
                List<Tuple<string, double, double>> ListGraphBarFini = new List<Tuple<string, double, double>>();
                var sommeRetraitJanvier = 0;
                var sommeRetraitFev = 0;
                var sommeRetraitMars = 0;
                var sommeRetraitAvril = 0;
                var sommeRetraitMai = 0;
                var sommeRetraitJuin = 0;
                var sommeRetraitJuillet = 0;
                var sommeRetraitAout = 0;
                var sommeRetraitSept = 0;
                var sommeRetraitOct = 0;
                var sommeRetraitNovembre = 0;
                var sommeRetraitDec = 0;

                var sommeVersementJanvier = 0;
                var sommeVersementFev = 0;
                var sommeVersementMars = 0;
                var sommeVersementAvril = 0;
                var sommeVersementMai = 0;
                var sommeVersementJuin = 0;
                var sommeVersementJuillet = 0;
                var sommeVersementAout = 0;
                var sommeVersementSept = 0;
                var sommeVersementOct = 0;
                var sommeVersementNovembre = 0;
                var sommeVersementDec = 0;

                if (avoir == null)
                {
                    foreach (var l in releveSituation.Avoirs)
                    {
                        sommeRetraitJanvier += l.SommeRetraitsMois(1);
                        sommeRetraitFev += l.SommeRetraitsMois(2);
                        sommeRetraitMars += l.SommeRetraitsMois(3);
                        sommeRetraitAvril += l.SommeRetraitsMois(4);
                        sommeRetraitMai += l.SommeRetraitsMois(5);
                        sommeRetraitJuin += l.SommeRetraitsMois(6);
                        sommeRetraitJuillet += l.SommeRetraitsMois(7);
                        sommeRetraitAout += l.SommeRetraitsMois(8);
                        sommeRetraitSept += l.SommeRetraitsMois(9);
                        sommeRetraitOct += l.SommeRetraitsMois(10);
                        sommeRetraitNovembre += l.SommeRetraitsMois(11);
                        sommeRetraitDec += l.SommeRetraitsMois(12);
                        // Somme Versement par mois :
                        sommeVersementJanvier += l.SommeVersementsMois(1);
                        sommeVersementFev += l.SommeVersementsMois(2);
                        sommeVersementMars += l.SommeVersementsMois(3);
                        sommeVersementAvril += l.SommeVersementsMois(4);
                        sommeVersementMai += l.SommeVersementsMois(5);
                        sommeVersementJuin += l.SommeVersementsMois(6);
                        sommeVersementJuillet += l.SommeVersementsMois(7);
                        sommeVersementAout += l.SommeVersementsMois(8);
                        sommeVersementSept += l.SommeVersementsMois(9);
                        sommeVersementOct += l.SommeVersementsMois(10);
                        sommeVersementNovembre += l.SommeVersementsMois(11);
                        sommeVersementDec += l.SommeVersementsMois(12);
                    }
                }
                else
                {
                    sommeRetraitJanvier += avoir.SommeRetraitsMois(1);
                    sommeRetraitFev += avoir.SommeRetraitsMois(2);
                    sommeRetraitMars += avoir.SommeRetraitsMois(3);
                    sommeRetraitAvril += avoir.SommeRetraitsMois(4);
                    sommeRetraitMai += avoir.SommeRetraitsMois(5);
                    sommeRetraitJuin += avoir.SommeRetraitsMois(6);
                    sommeRetraitJuillet += avoir.SommeRetraitsMois(7);
                    sommeRetraitAout += avoir.SommeRetraitsMois(8);
                    sommeRetraitSept += avoir.SommeRetraitsMois(9);
                    sommeRetraitOct += avoir.SommeRetraitsMois(10);
                    sommeRetraitNovembre += avoir.SommeRetraitsMois(11);
                    sommeRetraitDec += avoir.SommeRetraitsMois(12);
                    // Somme Versement par mois :
                    sommeVersementJanvier += avoir.SommeVersementsMois(1);
                    sommeVersementFev += avoir.SommeVersementsMois(2);
                    sommeVersementMars += avoir.SommeVersementsMois(3);
                    sommeVersementAvril += avoir.SommeVersementsMois(4);
                    sommeVersementMai += avoir.SommeVersementsMois(5);
                    sommeVersementJuin += avoir.SommeVersementsMois(6);
                    sommeVersementJuillet += avoir.SommeVersementsMois(7);
                    sommeVersementAout += avoir.SommeVersementsMois(8);
                    sommeVersementSept += avoir.SommeVersementsMois(9);
                    sommeVersementOct += avoir.SommeVersementsMois(10);
                    sommeVersementNovembre += avoir.SommeVersementsMois(11);
                    sommeVersementDec += avoir.SommeVersementsMois(12);
                }


                var MonthDebut = releveSituation.PeriodeDebut.Month;
                var MonthFin = releveSituation.PeriodeFin.Month;
                var SameYear = releveSituation.PeriodeDebut.Year == releveSituation.PeriodeFin.Year ? true : false;


                ListGraphBar.Add(new Tuple<string, double, double>("Jan", sommeVersementJanvier, sommeRetraitJanvier));
                ListGraphBar.Add(new Tuple<string, double, double>("Fev", sommeVersementFev, sommeRetraitFev));
                ListGraphBar.Add(new Tuple<string, double, double>("Mar", sommeVersementMars, sommeRetraitMars));
                ListGraphBar.Add(new Tuple<string, double, double>("Avr", sommeVersementAvril, sommeRetraitAvril));
                ListGraphBar.Add(new Tuple<string, double, double>("Mai", sommeVersementMai, sommeRetraitMai));
                ListGraphBar.Add(new Tuple<string, double, double>("Jun", sommeVersementJuin, sommeRetraitJuin));
                ListGraphBar.Add(new Tuple<string, double, double>("Jul", sommeVersementJuillet, sommeRetraitJuillet));
                ListGraphBar.Add(new Tuple<string, double, double>("Aou", sommeVersementAout, sommeRetraitAout));
                ListGraphBar.Add(new Tuple<string, double, double>("Sep", sommeVersementSept, sommeRetraitSept));
                ListGraphBar.Add(new Tuple<string, double, double>("Oct", sommeVersementOct, sommeRetraitOct));
                ListGraphBar.Add(new Tuple<string, double, double>("Nvb", sommeVersementNovembre, sommeRetraitNovembre));
                ListGraphBar.Add(new Tuple<string, double, double>("Dec", sommeVersementDec, sommeRetraitDec));

                var TempMothFin = MonthFin;
                if (!SameYear)
                {
                    TempMothFin = 12;
                }
                var count = 1;
                foreach (var y in ListGraphBar)
                {
                    if (count >= MonthDebut && count <= TempMothFin)
                    {
                        ListGraphBarFini.Add(y);
                    }
                    count++;
                }

                if (!SameYear)
                {
                    var count2 = 1;
                    foreach (var s in ListGraphBar)
                    {
                        if (count2 < MonthFin)
                        {
                            ListGraphBarFini.Add(s);
                        }
                        count2++;
                    }
                }
                return ListGraphBarFini;
            }

            catch (Exception e)
            {

                throw;
            }

        }

        private static List<Tuple<double[], string>> GetListEvolutionPatrimoine(ReleveSituation releveSituation)
        {
            List<Tuple<double[], string>> LineGraph = new List<Tuple<double[], string>>();
            try
            {
                var Element = releveSituation.Avoirs.Count > 8 ? releveSituation.Avoirs.OrderByDescending(x => x.Valorisation).Take(8) : releveSituation.Avoirs; // get maximum 8 elemnt 

                foreach (var e in Element)
                {
                    var NbElement = e.HistoriqueValorisations.Where(h => h.Key.Date >= releveSituation.PeriodeDebut.Date && h.Key.Date <= releveSituation.PeriodeFin.Date);
                    var Points = new double[NbElement.Count()];

                    var c = 0;
                    foreach (var d in e.HistoriqueValorisations.OrderBy(x => x.Key))
                    {
                        if (d.Key.Date >= releveSituation.PeriodeDebut.Date && d.Key.Date <= releveSituation.PeriodeFin.Date) // La periode doit etre dans la periode global 
                        {
                            Points[c] = d.Value;
                        }
                        c++;
                    }
                    LineGraph.Add(Tuple.Create(Points, e.Nom));
                }

                var SommeNbElement = releveSituation.GlobalHistoriqueValorisations.Where(s => s.Key >= releveSituation.PeriodeDebut.Date && s.Key.Date <= releveSituation.PeriodeFin.Date);
                var somme = new double[SommeNbElement.Count()];
                var countSomme = 0;
                foreach (var s in releveSituation.GlobalHistoriqueValorisations.OrderBy(x => x.Key))
                {
                    if (s.Key.Date >= releveSituation.PeriodeDebut.Date && s.Key.Date <= releveSituation.PeriodeFin.Date)
                    {
                        somme[countSomme] = s.Value;
                    }
                    countSomme++;
                }
                LineGraph.Insert(0, Tuple.Create(somme, "Totale"));

            }
            catch (Exception e)
            {
                throw;
            }

            return LineGraph;
        }

        private static List<string> GetMonthList(ReleveSituation releveSituation)
        {
            try
            {
                List<string> ListMonth = new List<string>();
                List<string> ListMonthFini = new List<string>();

                var MonthDebut = releveSituation.PeriodeDebut.Month;
                var MonthFin = releveSituation.PeriodeFin.Month;
                var SameYear = releveSituation.PeriodeDebut.Year == releveSituation.PeriodeFin.Year ? true : false;

                ListMonth.Add("Jan");
                ListMonth.Add("Fev");
                ListMonth.Add("Mar");
                ListMonth.Add("Avr");
                ListMonth.Add("Mai");
                ListMonth.Add("Jun");
                ListMonth.Add("Jul");
                ListMonth.Add("Aou");
                ListMonth.Add("Sep");
                ListMonth.Add("Oct");
                ListMonth.Add("Nvb");
                ListMonth.Add("Dec");

                var TempMothFin = MonthFin;
                if (!SameYear)
                {
                    TempMothFin = 12;
                }
                var count = 1;
                foreach (var y in ListMonth)
                {
                    if (count >= MonthDebut && count <= TempMothFin)
                    {
                        ListMonthFini.Add(y);
                    }
                    count++;
                }

                if (!SameYear)
                {
                    var count2 = 1;
                    foreach (var s in ListMonth)
                    {
                        if (count2 < MonthFin)
                        {
                            ListMonthFini.Add(s);
                        }
                        count2++;
                    }
                }
                return ListMonthFini;
            }
            catch (Exception e)
            {
                throw e;
            }

        }
        private static string GetTableOfRisque(double risque, string Couleur_Base_HEXA, string Td1WidhtParticulier = "width:55px; ", string OtherTdParticulier = "width:40px; ", string TableParticulier = "")
        {
            try
            {
                var TD1 = Td1WidhtParticulier;
                var TD2Et3 = OtherTdParticulier;
                var TD4Et5 = OtherTdParticulier;
                var TD6Et7 = OtherTdParticulier;

                var TableCss = "background-color:#efeded; text-align: center; font-size:12px;";

                if (risque < 2)
                {
                    TD1 += $"color:white; background-color:{Couleur_Base_HEXA};";
                }
                else if (risque < 4)
                {
                    TD2Et3 += $"color:white; background-color:{Couleur_Base_HEXA};";
                }
                else if (risque < 6)
                {
                    TD4Et5 += $"color:white; background-color:{Couleur_Base_HEXA};";
                }
                else
                {
                    TD6Et7 += $"color:white; background-color:{Couleur_Base_HEXA};";
                }

                var result = $"<table style='{TableCss}' {TableParticulier} >" +
                             $"<tr style='height:30px;'>" +
                             $"<td style='{TD1}' class='tdRisqueBorderWhite'> 1 </td>" +
                             $"<td style='{TD2Et3}' class='tdRisqueBorderWhite'> 2 </td> " +
                             $"<td style='{TD2Et3}' class='tdRisqueBorderWhite'> 3 </td> " +
                             $"<td style='{TD4Et5}' class='tdRisqueBorderWhite'> 4 </td> " +
                             $"<td style='{TD4Et5}' class='tdRisqueBorderWhite'> 5 </td>  " +
                             $"<td style='{TD6Et7}' class='tdRisqueBorderWhite'> 6 </td> " +
                             $"<td style='{TD6Et7}' class='tdRisqueBorderWhite'> 7 </td></tr> " +
                             $"<tr>" +
                             $"<td style='{TD1}' class='tdRisqueBorderWhite'>Très <br/> prudent</td> " +
                             $"<td colspan='2' style='{TD2Et3}' class='tdRisqueBorderWhite'>Prudent</td>  " +
                             $"<td colspan='2' style='{TD4Et5}' class='tdRisqueBorderWhite'>Equilibré</td> " +
                             $"<td colspan='2' style='{TD6Et7}' class='tdRisqueBorderWhite'>Dynamique</td>" +
                             $"</tr></table>";

                return result;
            }
            catch (Exception e)
            {

                throw e;
            }

        }

        private static Dictionary<string, double> GetDictionnarySortedBy4(Dictionary<string, double> dictionnaryBase)
        {
            try
            {
                var Result = new Dictionary<string, double>();
                var dictSorted = dictionnaryBase.OrderByDescending(p => p.Value);
                int count = 0;
                double autreSupport = 0;
                foreach (var i in dictSorted)
                {
                    if (count < 3)
                    {
                        Result.Add(i.Key, i.Value);
                    }
                    else
                    {
                        autreSupport += i.Value;
                    }
                    count++;
                }
                if (autreSupport > 0)
                {
                    Result.Add("Autres", autreSupport);
                }
                return Result;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        private static string GetGraphCamembert(Dictionary<string, double> Dict, System.Drawing.Color Couleur_Base_Drawing)
        {
            try
            {
                var path = Path.Combine(Directory, Guid.NewGuid() + ".png");
                using (var ch = new Chart())
                {
                    ch.ChartAreas.Add(new ChartArea());
                    ch.Palette = ChartColorPalette.None;
                    ch.PaletteCustomColors = GetColorLuminosite(Couleur_Base_Drawing).ToArray();
                    ch.BackColor = System.Drawing.Color.FromArgb(247, 247, 247);
                    ch.Legends.Add(new Legend());
                    ch.Legends[0].BackColor = System.Drawing.Color.FromArgb(247, 247, 247);

                    ch.Legends[0].Position.Auto = false;
                    ch.Legends[0].Position = new ElementPosition();
                    ch.Legends[0].Docking = Docking.Right;

                    ch.ChartAreas[0].BackColor = System.Drawing.Color.FromArgb(247, 247, 247);
                    ch.Height = 140;
                    ch.Width = 260;

                    var s = new Series
                    {
                        ChartType = SeriesChartType.Pie,
                        CustomProperties = "PieStartAngle=270,PieLineColor=Black",
                    };
                    foreach (var pnt in Dict)
                    {
                        s.Points.AddXY(pnt.Value == 0 ? "" : "", pnt.Value);
                    }

                    s.Points.Where(x => x.YValues[0] == 0).ToList().ForEach(x => x.CustomProperties = "PieLineColor = White");
                    ch.Series.Add(s);
                    ch.SaveImage(path, ChartImageFormat.Png);
                }
                return path;

            }
            catch (Exception e)
            {
                throw;
            }
        }

        private static string GetGraphBar(List<Tuple<string, double, double>> Dict, System.Drawing.Color hexaColorBase, int HeightGraphe = 180)
        {
            try
            {
                var path = Path.Combine(Directory, Guid.NewGuid() + ".png");
                using (var ch = new Chart())
                {
                    ch.ChartAreas.Add(new ChartArea());
                    ch.Palette = ChartColorPalette.None;
                    ch.PaletteCustomColors = GetColorLuminosite(hexaColorBase).ToArray();
                    ch.BackColor = System.Drawing.Color.FromArgb(247, 247, 247);

                    ch.Height = HeightGraphe;
                    ch.Width = 700;
                    ch.ChartAreas[0].AxisY.Maximum = 480; // sets the Maximum to NaN

                    ch.ChartAreas[0].BackColor = System.Drawing.Color.FromArgb(247, 247, 247);
                    ch.ChartAreas[0].BorderWidth = 0;
                    ch.ChartAreas[0].BorderDashStyle = ChartDashStyle.NotSet;
                    ch.ChartAreas[0].RecalculateAxesScale();

                    ch.ChartAreas[0].AxisX.MajorGrid.LineWidth = 0;

                    ch.ChartAreas[0].AxisX.MajorTickMark.Enabled = false;// supprimer les tiret de l'axe X
                    ch.ChartAreas[0].AxisX.MinorTickMark.Enabled = false;// supprimer les tiret de l'axe X

                    ch.ChartAreas[0].AxisY.MajorTickMark.Enabled = false;// supprimer les tiret de l'axe Y
                    ch.ChartAreas[0].AxisY.MinorTickMark.Enabled = false;// supprimer les tiret de l'axe Y

                    double init = 0;
                    double end = 2.8;
                    foreach (var D in Dict)
                    {
                        ch.ChartAreas[0].AxisX.CustomLabels.Add(init, end, D.Item1);
                        init = end + 0.1;
                        end = end + 2;
                    }

                    ch.ChartAreas[0].AxisY.MajorGrid.LineColor = System.Drawing.Color.FromArgb(127, 140, 141);

                    var s = new Series
                    {
                        ChartType = SeriesChartType.Column,
                        CustomProperties = "PieStartAngle=270,PieLineColor=Black"
                    };
                    foreach (var pnt in Dict)
                    {
                        s.Points.AddY(pnt.Item2);
                        s.Points.AddY(pnt.Item3);
                    }
                    foreach (var p in s.Points)
                    {
                        if (p.YValues[0] < 0)
                            p.Color = System.Drawing.Color.FromArgb(4, 35, 53);
                        else
                            p.Color = hexaColorBase;
                    }
                    s.Points.Where(x => x.YValues[0] == 0).ToList().ForEach(x => x.CustomProperties = "PieLineColor = White");
                    ch.Series.Add(s);
                    ch.SaveImage(path, ChartImageFormat.Png);
                }
                return path;

            }
            catch (Exception e)
            {
                throw e;
            }
        }

        private static string GetGraphLine(List<Tuple<double[], string>> ListPoints, ReleveSituation releveSituation, System.Drawing.Color couleurBase, int HeightGraphe = 200, bool WithTotalLine = false)
        {
            try
            {
                var path = Path.Combine(Directory, Guid.NewGuid() + ".png");
                using (var ch = new Chart())
                {
                    ch.ChartAreas.Add(new ChartArea());
                    ch.Palette = ChartColorPalette.None;
                    List<System.Drawing.Color> ListColor = GetColorLuminosite(couleurBase, WithTotalLine);
                    ch.PaletteCustomColors = ListColor.ToArray();
                    ch.BackColor = System.Drawing.Color.FromArgb(247, 247, 247);

                    ch.Legends.Add(new Legend());
                    ch.Legends[0].BackColor = System.Drawing.Color.FromArgb(236, 236, 236);
                    ch.Legends[0].LegendStyle = LegendStyle.Table; // style legende tableau 

                    //ch.Legends[0].Position.Auto = false;
                    //ch.Legends[0].Position = new ElementPosition(30, 85, 100, 20);
                    ch.Legends[0].Docking = Docking.Bottom; // Positioner les legendes en bas du graphe 


                    ch.Height = HeightGraphe;
                    ch.Width = 700;
                    //ch.ChartAreas[0].AxisY.Maximum = 480/*Double.NaN*/; // sets the Maximum to NaN

                    ch.ChartAreas[0].BackColor = System.Drawing.Color.FromArgb(247, 247, 247);
                    ch.ChartAreas[0].BorderWidth = 0;
                    ch.ChartAreas[0].BorderDashStyle = ChartDashStyle.NotSet;
                    ch.ChartAreas[0].RecalculateAxesScale();

                    ch.ChartAreas[0].AxisX.MajorGrid.LineWidth = 0;
                    ch.ChartAreas[0].AxisY.MajorGrid.LineColor = System.Drawing.Color.FromArgb(127, 140, 141);

                    List<Series> AllCourbe = new List<Series>();
                    //GET LABEL NAME : 
                    List<string> ListLabel = GetMonthList(releveSituation);
                    var countColor = 0;
                    foreach (var l in ListPoints)
                    {
                        var s = new Series
                        {
                            Name = l.Item2,
                            ChartType = SeriesChartType.Line,
                            CustomProperties = "PieStartAngle=270,PieLineColor=Black",
                            Color = ListColor[countColor],
                        };
                        countColor++;
                        foreach (var pnt in l.Item1)
                        {
                            s.Points.AddXY("", pnt); //add y en principe 
                        }
                        AllCourbe.Add(s);
                    }

                    var count = 0;
                    var countLabel = 1;
                    foreach (var l in AllCourbe) // liste des courbes
                    {
                        foreach (var pnt in l.Points)
                        {
                            pnt.Color = ListColor[count]; // colorer les lignes de graphes 

                            pnt.AxisLabel = ListLabel[countLabel - 1]; // Label 
                            pnt.MarkerStyle = MarkerStyle.Circle; // Style Point Marker 
                            countLabel++;
                        }
                        countLabel -= l.Points.Count;
                        l.Points.Where(x => x.YValues[0] == 0).ToList().ForEach(x => x.CustomProperties = "PieLineColor = White");

                        ch.Series.Add(l);
                        count++;
                    }

                    ch.SaveImage(path, ChartImageFormat.Png);
                }
                return path;
            }
            catch (Exception e)
            {
                throw;
            }
        }


        public static void AddHeaderFooterPDF(string PathFile, string outFile, BaseColor color, ReleveSituation releveSituation)
        {
            using (var reader = new PdfReader(PathFile))
            using (var fileStream = new FileStream(Path.Combine(Directory, outFile), FileMode.Create, FileAccess.Write))
            {
                try
                {
                    BaseColor fontColourBlackGris = new BaseColor(65, 64, 66);
                    Font FontWhitePetiteBold = FontFactory.GetFont("Calibri", 12, Font.BOLD, releveSituation.CouleurEcriture);
                    Font FontWhitePetite = FontFactory.GetFont("Calibri", 12, releveSituation.CouleurEcriture);
                    Font FontBlackGrisPetite = FontFactory.GetFont("Calibri", 8, fontColourBlackGris);
                    using (var document = new Document(reader.GetPageSizeWithRotation(1)))
                    using (var stamper = new PdfStamper(reader, fileStream))
                    {
                        document.Open();
                        var firstparag = $"Relevé de situation de {releveSituation.Civilite}. {releveSituation.Prenom} {releveSituation.Nom}";
                        var secondparag = $"Du {releveSituation.PeriodeDebut.ToString("dd/MM/yyy")} {releveSituation.PeriodeFin.ToString("dd/MM/yyy")}";
                        for (var i = 1; i <= reader.NumberOfPages; i++)
                        {
                            PdfContentByte CbForAddingHeader = stamper.GetOverContent(i);
                            CbForAddingHeader.Rectangle(0, PageSize.A4.Height - 85, PageSize.A4.Width - 170, 85);

                            CbForAddingHeader.MoveTo(PageSize.A4.Width - 170, PageSize.A4.Height);
                            CbForAddingHeader.LineTo(PageSize.A4.Width - 100, PageSize.A4.Height);
                            CbForAddingHeader.LineTo(PageSize.A4.Width - 170, PageSize.A4.Height - 85);

                            // Triangle for Number of pages :
                            CbForAddingHeader.MoveTo(PageSize.A4.Width - 100, 0);
                            CbForAddingHeader.LineTo(PageSize.A4.Width, 0);
                            CbForAddingHeader.LineTo(PageSize.A4.Width, 85);


                            CbForAddingHeader.SetColorFill(color);
                            CbForAddingHeader.Fill();
                            CbForAddingHeader.SetColorStroke(color);
                            CbForAddingHeader.FillStroke();
                            CbForAddingHeader.Stroke();
                            CbForAddingHeader.ClosePathStroke();

                            ColumnText.ShowTextAligned(CbForAddingHeader, Element.ALIGN_LEFT, new Phrase((i + 1).ToString(), FontWhitePetite), PageSize.A4.Width - 30, 25, 0); // Number of page
                            ColumnText.ShowTextAligned(CbForAddingHeader, Element.ALIGN_LEFT, new Phrase(firstparag, FontWhitePetiteBold), 30, PageSize.A4.Height - 30, 0);
                            ColumnText.ShowTextAligned(CbForAddingHeader, Element.ALIGN_LEFT, new Phrase(secondparag, FontWhitePetite), 30, PageSize.A4.Height - 45, 0);

                            //Pieds de Page : 
                            ColumnText.ShowTextAligned(CbForAddingHeader, Element.ALIGN_LEFT, new Phrase(releveSituation.Footer1, FontBlackGrisPetite), 30, 50, 0);
                            ColumnText.ShowTextAligned(CbForAddingHeader, Element.ALIGN_LEFT, new Phrase(releveSituation.Footer2, FontBlackGrisPetite), 30, 40, 0);
                            ColumnText.ShowTextAligned(CbForAddingHeader, Element.ALIGN_LEFT, new Phrase(releveSituation.Footer3, FontBlackGrisPetite), 30, 30, 0);
                            // Logo
                            Image Logo = Image.GetInstance(releveSituation.PathLogo);
                            Logo.SetAbsolutePosition(350, 30);
                            Logo.ScaleAbsoluteHeight(Logo.Height * 0.5f);
                            Logo.ScaleAbsoluteWidth(Logo.Width * 0.5f);
                            CbForAddingHeader.AddImage(Logo);


                        }
                    }


                }
                catch (Exception e)
                {
                    throw;
                }
            }
        }

        public static bool MergePdfInPosition(string pathNormalPdf, string pdfToAdded, string outFile, int positionMerge)
        {
            try
            {
                using (FileStream stream = new FileStream(outFile, FileMode.Append))
                {
                    using (iTextSharp.text.Document doc = new iTextSharp.text.Document())
                    {
                        using (PdfCopy pdf = new PdfCopy(doc, stream))
                        {
                            doc.Open();
                            PdfImportedPage page = null;

                            var readerInitDoc = new PdfReader(pathNormalPdf);
                            var readerSecondDoc = new PdfReader(pdfToAdded);
                            var MergedIsDone = false;
                            for (int i = 0; i < readerInitDoc.NumberOfPages; i++)
                            {
                                if (i != positionMerge || MergedIsDone)
                                {
                                    page = pdf.GetImportedPage(readerInitDoc, i + 1);
                                    pdf.AddPage(page);
                                }
                                else
                                {
                                    int j;
                                    for (j = 0; j < readerSecondDoc.NumberOfPages; j++)
                                    {
                                        page = pdf.GetImportedPage(readerSecondDoc, j + 1);
                                        pdf.AddPage(page);
                                    }
                                    i = i - 1;
                                    MergedIsDone = true;
                                }
                            }
                            readerSecondDoc.Close();
                            readerInitDoc.Close();
                            doc.Close();
                        }
                    }
                }
                return true;
            }
            catch (Exception e)
            {
                // # TODO add some log here

            }
            return false;
        }

        public static string BuildWorldMap(Dictionary<ZoneGeographique, double> RepartitionGeographique, string Couleur_Base_HEXA)
        {
            try
            {
                var AmeriqueSudColor = RepartitionGeographique.ContainsKey(ZoneGeographique.AmeriqueSud) ? Couleur_Base_HEXA : "#f9f9f9";
                var AmeriqueNordColor = RepartitionGeographique.ContainsKey(ZoneGeographique.AmeriqueNord) ? Couleur_Base_HEXA : "#f9f9f9";
                var EuropeColor = RepartitionGeographique.ContainsKey(ZoneGeographique.Europe) ? Couleur_Base_HEXA : "#f9f9f9";
                var AsiaColor = RepartitionGeographique.ContainsKey(ZoneGeographique.Asie) ? Couleur_Base_HEXA : "#f9f9f9";
                var AfriqueColor = RepartitionGeographique.ContainsKey(ZoneGeographique.Afrique) ? Couleur_Base_HEXA : "#f9f9f9";
                var OceanieColor = RepartitionGeographique.ContainsKey(ZoneGeographique.Oceanie) ? Couleur_Base_HEXA : "#f9f9f9";

                var AmeriqueSud = RepartitionGeographique.ContainsKey(ZoneGeographique.AmeriqueSud) ? "Amerique du sud " + RepartitionGeographique.Where(x => (int)x.Key == 1).FirstOrDefault().Value + "%" : "";
                var AmeriqueNord = RepartitionGeographique.ContainsKey(ZoneGeographique.AmeriqueNord) ? "Amerique du nord " + RepartitionGeographique.Where(x => (int)x.Key == 0).FirstOrDefault().Value + "%" : "";
                var Europe = RepartitionGeographique.ContainsKey(ZoneGeographique.Europe) ? "Europe " + RepartitionGeographique.Where(x => (int)x.Key == 2).FirstOrDefault().Value + "%" : "";
                var Afrique = RepartitionGeographique.ContainsKey(ZoneGeographique.Afrique) ? "Afrique " + RepartitionGeographique.Where(x => (int)x.Key == 3).FirstOrDefault().Value + "%" : "";
                var Asia = RepartitionGeographique.ContainsKey(ZoneGeographique.Asie) ? "Asia " + RepartitionGeographique.Where(x => (int)x.Key == 4).FirstOrDefault().Value + "%" : "";
                var Oceanie = RepartitionGeographique.ContainsKey(ZoneGeographique.Oceanie) ? "Oceanie " + RepartitionGeographique.Where(x => (int)x.Key == 5).FirstOrDefault().Value + "%" : "";
                var PaysEmergent = RepartitionGeographique.ContainsKey(ZoneGeographique.PaysEmergents) ? "Pays emergent " + RepartitionGeographique.Where(x => (int)x.Key == 6).FirstOrDefault().Value + "%" : "";

                var convertor = new CoreHtmlToImage.HtmlConverter();
                string HtmlToConvert = File.ReadAllText(Path.Combine(TemplatesDirectory, "map_template.html")).Replace("[Zone_Amerique_Sud]", AmeriqueSud).Replace("[Zone_Amerique_Nord]", AmeriqueNord).Replace("[Zone_Europe]", Europe).Replace("[Zone_Afrique]", Afrique).Replace("[Zone_Asia]", Asia).Replace("[Zone_Oceanie]", Oceanie).Replace("[Zone_PaysEmergent]", PaysEmergent);
                HtmlToConvert = HtmlToConvert.Replace("[Zone_Couleur_Amerique_Sud]", AmeriqueSudColor).Replace("[Zone_Couleur_Amerique_Nord]", AmeriqueNordColor).Replace("[Zone_Europe_Couleur]", EuropeColor).Replace("[Zone_Asia_Couleur]", AsiaColor).Replace("[Zone_Oceanie_Couleur]", OceanieColor).Replace("[Zone_Afrique_Couleur]", AfriqueColor);
                var templateBytes = convertor.FromHtmlString(HtmlToConvert, 500, CoreHtmlToImage.ImageFormat.Png, 100);
                var targetPngPath = Path.Combine(Directory, "graph.png");
                File.WriteAllBytes(targetPngPath, templateBytes);
                return targetPngPath;
            }
            catch (Exception e)
            {
                throw;
            }

        }

        private static List<System.Drawing.Color> GetColorLuminosite(System.Drawing.Color BaseColor, bool IsLineGraph = false)
        {
            var ListColor = new List<System.Drawing.Color>();
            try
            {
                if (IsLineGraph) // si c'est pour le line graphe il faut rajouter la couleur de Totale
                {
                    ListColor.Add(System.Drawing.Color.Black);
                }

                ListColor.Add(BaseColor); // couleur de base 
                double h, s, l;
                int r, g, b;

                ColorHelpers.RgbToHls(BaseColor.R, BaseColor.G, BaseColor.B, out h, out l, out s);
                var oldL = l;
                bool LMax = false;
                for (var a = 1; a < 9; a++)
                {
                    var tempL = l;
                    l += 0.16;
                    if (l > 1 || LMax)
                    {
                        LMax = true;
                        oldL = oldL - 0.16 < 0 ? 0 : oldL - 0.16;
                        l = oldL;
                    }

                    ColorHelpers.HlsToRgb(h, l, s, out r, out g, out b);
                    var color = System.Drawing.Color.FromArgb(r, g, b);
                    ListColor.Add(color);
                }
            }
            catch (Exception e)
            {
                throw;
            }

            return ListColor;
        }

        private static void InsertSommaire(string DocFinalReader, Font FontSommaire, List<Tuple<string, int>> Sommaire, ReleveSituation releveSituation, string PathPageDeGarde)
        {
            try
            {
                var PathFileWithSommaire = DocFinalReader.Replace(".pdf", $"WithSommaire.pdf{DateTime.Now:ddMMyyyy_HH-mm-ss}");
                var ColorShape = releveSituation.CouleurBase;
                using (var reader = new PdfReader(DocFinalReader))
                using (FileStream stream1 = File.Open(PathFileWithSommaire, FileMode.OpenOrCreate))
                using (PdfStamper stamper = new PdfStamper(reader, stream1))
                {

                    PdfContentByte Cb = stamper.GetOverContent(1);

                    // calcul how much page are added to sommaire for edit number page of each section in sommaire 
                    var GlobalNbPageAddedToSommaire = ((Sommaire.Count) / 27);
                    GlobalNbPageAddedToSommaire = ((Sommaire.Count) % 27 == 0) ? GlobalNbPageAddedToSommaire - 1 : GlobalNbPageAddedToSommaire;

                    ColumnText.ShowTextAligned(Cb, Element.ALIGN_LEFT, new Phrase("Vue globale", FontSommaire), 85, 670, 0);
                    ColumnText.ShowTextAligned(Cb, Element.ALIGN_LEFT, new Phrase((3 + GlobalNbPageAddedToSommaire).ToString(), FontSommaire), 480, 670, 0);

                    var pageAdded = 0;
                    var countLine = 0;
                    var y = 650;
                    foreach (var s in Sommaire)
                    {
                        countLine++;
                        if (countLine % 27 == 0)
                        {
                            stamper.InsertPage(2 + pageAdded, reader.GetPageSize(1));
                            y = 650;
                            Cb = stamper.GetOverContent(2 + pageAdded);
                            pageAdded++;
                        }

                        ColumnText.ShowTextAligned(Cb, Element.ALIGN_LEFT, new Phrase(s.Item1, FontSommaire), 85, y, 0);
                        ColumnText.ShowTextAligned(Cb, Element.ALIGN_LEFT, new Phrase((s.Item2 + GlobalNbPageAddedToSommaire).ToString(), FontSommaire), 480, y, 0);
                        y -= 20;
                    }
                }

                File.Delete(DocFinalReader);
                AddHeaderFooterPDF(Path.Combine(Directory, PathFileWithSommaire), "OutFile.pdf", ColorShape, releveSituation);
                MergePdfInPosition(Path.Combine(Directory, "OutFile.pdf"), Path.Combine(Directory, PathPageDeGarde), DocFinalReader, 0);
                File.Delete(PathFileWithSommaire);
                File.Delete(Path.Combine(Directory, PathPageDeGarde));
                File.Delete(Path.Combine(Directory, "OutFile.pdf"));

            }

            catch (Exception e)
            {

                throw;
            }
        }

        private static void InitObject(ReleveSituation releveSituation)
        {
            if (releveSituation.CouleurBase == null) // test si ça n'existe pas une couleurde base en parametre mette par defaut noir 
            {
                releveSituation.CouleurBase = new BaseColor(0, 0, 0);
            }

            if (releveSituation.CouleurEcriture == null) // test si ça n'existe pas une couleur d'ecriture en parametre mette par defaut noir 
            {
                releveSituation.CouleurEcriture = new BaseColor(255, 255, 255);
            }

            if (releveSituation.PeriodeDebut == null)
            {
                releveSituation.PeriodeDebut = new DateTime(1900, 01, 01);
            }

            if (releveSituation.PeriodeFin == null)
            {
                releveSituation.PeriodeFin = new DateTime(1900, 01, 01);
            }

            if (releveSituation.Avoirs == null)
            {
                releveSituation.Avoirs = new List<Avoir>();
            }

            if (releveSituation.Frais == null)
            {
                releveSituation.Frais = new List<Frais>();
            }

        }
    }
}
