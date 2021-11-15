using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Collections.ObjectModel;
using Excel = Microsoft.Office.Interop.Excel;
using System.ComponentModel;
//using System.Threading;

namespace HebdoJEL
{
    enum ColId{ Semaine_du = 1, Plateforme, Nom, Usage, Nb_de_jours_en_violations, Niveau_de_criticite, Date, Evenement }
    enum NbCat { CAT1, CAT2, CAT3, CAT4, NombreCAT = 4}
    /// <summary>
    /// Logique d'interaction pour App.xaml
    /// </summary>
    public partial class App : Application
    {
        private const string PATH_VÉRIFICATION_SYSTÈMES_JOURNAL = @"\\le500\dfs\INTERLQ\CASJPartage\Vérification\Système2\Fichiers résultats (journal qtd)\";
        //private const string PATH_VÉRIFICATION_SYSTÈMES_JOURNAL = @"\\le500\dfs\VP\CAJ\COMMUN\DCCASJ\Laboratoire DCCASJ\xDEV\JAB\Vérification\Système2\Fichiers résultats (journal qtd)\";
        private const string PATH_VÉRIFICATION_SYSTÈMES_JOURNAL_TRAITES = PATH_VÉRIFICATION_SYSTÈMES_JOURNAL + @"TRAITES\";
        private const string PATH_VÉRIFICATION_SYSTÈMES_JOURNAL_RAPPORT_HEBDO = PATH_VÉRIFICATION_SYSTÈMES_JOURNAL + @"Rapports Hebdo\";
        private const string PATH_VÉRIFICATION_INDICATEURS = @"\\le500\dfs\VP\CAJ\COMMUN\DCCASJ\Vérification\Indicateurs\";

        private static Excel.Application excelApp;
        private static Excel.Workbook excelWB;
        //private static Excel.Workbook[] excelWBx = new Excel.Workbook[7]
        private static Excel.Worksheet excelWS;
        private static Excel.Range xlRange;
        private static readonly int[] ligneCatxPrecedente = new int[(int)NbCat.NombreCAT]; // Nb de CAT

        //***************************************************************************************************
        //* DateValidation - Retourne une liste contenant les dates pour lesquelles les journaux quotidiens 
        //*                  n'ont pu être localisés
        //*
        //*                - dateDebut:  Date (DateTime) de début de la semaine désirée
        //*                - Retourne :  Une liste (string) contenant les fichiers manquants
        //*
        //***************************************************************************************************
        public static List<string> DateValidation(DateTime dateDebut)
        {
            IEnumerable<FileInfo> files;
            DateTime dateInterval;
            string dateStrInterval;
            int i = 0;
            List<string> filesM = new List<string>();
            List<string> FichierManquant = new List<string>();

            DirectoryInfo DirInfo = new DirectoryInfo(PATH_VÉRIFICATION_SYSTÈMES_JOURNAL);
            dateStrInterval = dateDebut.ToString("d");
            try
            {
                // Query for all files created after or at dateDebut.
                files = from f in DirInfo.GetFiles(dateStrInterval.Substring(0, 2) + "??-??-??_Journal_VériJEL.csv", SearchOption.TopDirectoryOnly)
                        where f.CreationTime >= dateDebut
                        select f;
            }
            catch (Exception ex)
            {
                FichierManquant.Add(ex.Message);
                return FichierManquant;
            }

            foreach (FileInfo F in files)
            {
                filesM.Add(F.Name);
            }

            while (i < 7)
            {
                dateInterval = dateDebut.AddDays(i);
                dateStrInterval = dateInterval.ToString("d");

                if (filesM.Contains(dateStrInterval + "_Journal_VériJEL.csv"))
                {
                    Debug.WriteLine(dateStrInterval + "_Journal_VériJEL.csv <EXIST>");
                    filesM.Remove(dateStrInterval + "_Journal_VériJEL.csv");
                }
                else
                {
                    Debug.WriteLine(dateStrInterval + "_Journal_VériJEL.csv <DON'T EXIST> ");
                    FichierManquant.Add(dateStrInterval);
                }
                i++;
            }
            Debug.WriteLine("DateValidation <EXIT>");
            return FichierManquant;
        }

        //private static BackgroundWorker BGW_Loading_Indicator;
        //private static BackgroundWorker BGW_Loading_Indicator = new BackgroundWorker
        //{
        //    WorkerReportsProgress = true,
        //    WorkerSupportsCancellation = true
        //};

        //public static bool Success = false;
        //private static DateTime dateDebut = DateTime.MinValue;

        ////***************************************************************************************************
        ////* BGW_Loading_Indicator_DoWork - 
        ////*
        ////*              - sender:  
        ////*              - e: 
        ////*
        ////***************************************************************************************************
        //static void BGW_Loading_Indicator_DoWork(object sender, DoWorkEventArgs e)
        //{

        //    BGW_Loading_Indicator.ProgressChanged += BGW_Loading_Indicator_ProgressChanged;
        //    //Do the long running process
        //    App.Success = JournalHebdo(App.dateDebut);
        //}

        ////***************************************************************************************************
        ////* BGW_Loading_Indicator_ProgressChanged - 
        ////*
        ////*              - sender:  
        ////*              - e: 
        ////*
        ////***************************************************************************************************
        //static void BGW_Loading_Indicator_ProgressChanged(object sender, ProgressChangedEventArgs e)
        //{
        //    Console.WriteLine("Completed" + e.ProgressPercentage + "%");
        //}

        ////***************************************************************************************************
        ////* BGW_Loading_Indicator_RunWorkerCompleted - 
        ////*
        ////*              - sender:  
        ////*              - e: 
        ////*
        ////***************************************************************************************************
        //static void BGW_Loading_Indicator_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        //{
        //    //Hide your wait dialog
        //    //CreationFait = true;
        //    if (e.Cancelled)
        //    {
        //        Console.WriteLine("Operation Cancelled");
        //    }
        //    else if (e.Error != null)
        //    {
        //        Console.WriteLine("Error in Process :" + e.Error);
        //    }
        //    else
        //    {
        //        Console.WriteLine("Operation Completed :" + e.Result);
        //    }
        //}

        ////***************************************************************************************************
        ////* MainJournalHebdo - Départ des tâches pour la création du journal hebdomadaire en multi-tâches.
        ////*
        ////*              - dateDebut:  Date (DateTime) de début de la semaine désirée
        ////*              - Retourne:  true pour succès et false pour un fail.
        ////*
        ////***************************************************************************************************
        //public static bool MainJournalHebdo(DateTime dateDeb)
        //{
        //    //    Show your wait dialog
        //    //BackgroundWorker BGW_Loading_Indicator = new BackgroundWorker
        //    //{
        //    //    WorkerReportsProgress = true,
        //    //    WorkerSupportsCancellation = true
        //    //};
        //    //CreationFait = false;
        //    // Subscribe to events
        //    BGW_Loading_Indicator.DoWork += BGW_Loading_Indicator_DoWork;
        //    BGW_Loading_Indicator.RunWorkerCompleted += BGW_Loading_Indicator_RunWorkerCompleted;
        //    //BGW_Loading_Indicator.ProgressChanged += BGW_Loading_Indicator_ProgressChanged;

        //    App.dateDebut = dateDeb;
        //    BGW_Loading_Indicator.RunWorkerAsync();
        //    return Success;
        //}

        //***************************************************************************************************
        //* JournalHebdo - Mise en place des liste des rapports journaliers pour la création du journal 
        //*                hebdomadaire ayant pour date de départ dateDebut suivi de l'appel de la fonction 
        //*                principale (FormatJournalHebdo)
        //*
        //*              - dateDebut:  Date (DateTime) de début de la semaine désirée
        //*              - Retourne:  true pour succès et false pour un fail.
        //*
        //***************************************************************************************************
        public static bool JournalHebdo(DateTime dateDepart)
        {
            DateTime dateInterval;
            IEnumerable<FileInfo> files, filesTr;
            List<string> filesHebdo = new List<string>();
            List<string> filesHebdoTr = new List<string>();
            string dateStrInterval;
            List<string> FichierToHebdo = new List<string>();
            List<string> FichierToHebdoTr = new List<string>();

            DirectoryInfo DirFichRes = new DirectoryInfo(PATH_VÉRIFICATION_SYSTÈMES_JOURNAL);
            DirectoryInfo DirJourTraites = new DirectoryInfo(PATH_VÉRIFICATION_SYSTÈMES_JOURNAL_TRAITES);
            dateStrInterval = dateDepart.ToString("d");
            int i = 0;

            try
            {
                // Query for all files created after or at dateDebut.
                files = from f in DirFichRes.GetFiles(dateStrInterval.Substring(0, 2) + "??-??-??_Journal_VériJEL.csv", SearchOption.TopDirectoryOnly)
                        where DateTime.Parse(f.Name.Substring(0, 10)) >= DateTime.Parse(dateStrInterval.Substring(0, 10))
                        select f;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                //BGW Success = false;
                return false;
            }

            try
            {
                // Query for all files created after or at dateDebut.
                filesTr = from f in DirJourTraites.GetFiles(dateStrInterval.Substring(0, 2) + "??-??-??_Journal_VériJEL.csv", SearchOption.TopDirectoryOnly)
                          where DateTime.Parse(f.Name.Substring(0, 10)) >= DateTime.Parse(dateStrInterval.Substring(0, 10))
                          select f;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }

            foreach (FileInfo F in files)
            {
                filesHebdo.Add(F.Name);
            }

            foreach (FileInfo F in filesTr)
            {
                filesHebdoTr.Add(F.Name);
            }

            do
            {
                dateInterval = dateDepart.AddDays(i);
                dateStrInterval = dateInterval.ToString("d");

                if (filesHebdoTr.Contains(dateStrInterval + "_Journal_VériJEL.csv"))
                {
                    Debug.WriteLine("TRAITES: " + dateStrInterval + "_Journal_VériJEL.csv <EXIST>");
                    filesHebdoTr.Remove(dateStrInterval + "_Journal_VériJEL.csv");
                    FichierToHebdoTr.Add(dateStrInterval + "_Journal_VériJEL.csv");
                }
                else
                {
                    Debug.WriteLine("NON_T: " + dateStrInterval + "_Journal_VériJEL.csv <EXIST>");
                    filesHebdo.Remove(dateStrInterval + "_Journal_VériJEL.csv");
                    FichierToHebdo.Add(dateStrInterval + "_Journal_VériJEL.csv");
                }
                i++;
            } while (i < 7);

            FormatJournalHebdo(dateDepart, FichierToHebdo, FichierToHebdoTr);

            return true;
        }

        //***************************************************************************************************
        //* FormatJournalHebdo - 
        //*
        //*                    - dateDebut:  Date (DateTime) de début de la semaine désirée
        //*                    - FichierToHebdo: Liste (string) des fichiers journaliers existants
        //*                    - FichierToHebdoTr: Liste (string) des fichiers journaliers traités existants
        //*
        //***************************************************************************************************
        private static void FormatJournalHebdo(DateTime dateDebut, List<string> FichierToHebdo, List<string> FichierToHebdoTr)
        {

            CreationFichierJournal(dateDebut);
            ConcatFichiers(dateDebut, FichierToHebdo, FichierToHebdoTr);
            FormatFichier(dateDebut);
            PivotMimic();
            CleanupJournal();
            FermetureExcel(PATH_VÉRIFICATION_SYSTÈMES_JOURNAL_RAPPORT_HEBDO + dateDebut.ToShortDateString() + "_Journal_Hebdo.csv"); // dateDebut);
            AppendToFile(dateDebut, "Anomalies Verification JEL.csv");
            FermetureExcel(PATH_VÉRIFICATION_INDICATEURS + "Anomalies Verification JEL.csv");
        }

        //***************************************************************************************************
        //* AppendToFile - 
        //*
        //*                    - dateDebut : Date du début du journal contenant des anomalies à ajouter
        //*                    - fichierAnomalies:  Nom du fichier auquel ajouter les nouvelles anomalies 
        //*                      observées
        //*
        //***************************************************************************************************
        private static void AppendToFile(DateTime dateDebut, string fichierAnomalies)
        {

            string fichierAno, nomJournal;
            int lignesCompteur, colonnesCompteur;
            Excel.Range entete, ligneDel;

            fichierAno = PATH_VÉRIFICATION_INDICATEURS + fichierAnomalies;
            nomJournal = PATH_VÉRIFICATION_SYSTÈMES_JOURNAL_RAPPORT_HEBDO + dateDebut.ToShortDateString() + "_Journal_Hebdo.csv";

            File.AppendAllText(fichierAno, File.ReadAllText(nomJournal.Trim()));

            excelApp = new Excel.Application();
            {
                var withBlock = excelApp;
#if DEBUG
                withBlock.Visible = true;                       // Ajuster à False pour prod
                withBlock.Application.ScreenUpdating = true;    // Ajuster à False pour prod
#else
                withBlock.Visible = false;                      // Ajuster à False pour prod
                withBlock.Application.ScreenUpdating = false;   // Ajuster à False pour prod
#endif
                withBlock.Application.DisplayAlerts = false;
                withBlock.Application.EnableEvents = false;
            }
            // OUVERTURE EXCEL (Local:= True)
            excelWB = excelApp.Workbooks.Open(Filename: fichierAno, UpdateLinks: 2, ReadOnly: false, AddToMru: false, Notify: false, Local: true);
            excelWS = excelWB.ActiveSheet;

            lignesCompteur = excelWS.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            // Find the last real column
            colonnesCompteur = excelWS.Cells.Find("*", System.Reflection.Missing.Value,
                                           System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                           Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious,
                                           false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;

            entete = excelWS.Range[excelWS.Cells[2, 1], excelWS.Cells[lignesCompteur, colonnesCompteur]];

            /***** Effacer la ligne d'entête *****/
            // You should specify all these parameters every time you call this method,
            // since they can be overridden in the user interface. 
            entete = entete.Find(excelWS.Cells[1, 1], System.Reflection.Missing.Value, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                                          Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, System.Reflection.Missing.Value,
                                          System.Reflection.Missing.Value);

            ligneDel = excelWS.Rows[entete.Row];
            ligneDel.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
        }

        //***************************************************************************************************
        //* CreationFichierJournal - Création du fichier XXX_Journal_Hebdo.csv avec l'entête
        //*                  
        //*                        - dateDebut:  Date (DateTime) de début de la semaine désirée
        //*
        //***************************************************************************************************
        private static void CreationFichierJournal(DateTime dateDebut)
        {
            string nomJournal;
            object misValue = System.Reflection.Missing.Value;

            excelApp = new Excel.Application();
            {
                var withBlock = excelApp;
#if DEBUG
                withBlock.Visible = true;                       // Ajuster à False pour prod
                withBlock.Application.ScreenUpdating = true;    // Ajuster à False pour prod
#else
                withBlock.Visible = false;                      // Ajuster à False pour prod
                withBlock.Application.ScreenUpdating = false;   // Ajuster à False pour prod
#endif
                withBlock.Application.DisplayAlerts = false;
                withBlock.Application.EnableEvents = false;
            }

            excelWB = excelApp.Workbooks.Add(misValue);
            excelWS = excelWB.Worksheets.get_Item(1);

            excelWS.Cells[1, 1] = "Semaine du ";
            excelWS.Cells[1, 2] = "Plateforme";
            excelWS.Cells[1, 3] = "Nom";
            excelWS.Cells[1, 4] = "Usage";
            excelWS.Cells[1, 5] = "Nb de jours en violations";
            excelWS.Cells[1, 6] = "Niveau de criticité";

            nomJournal = PATH_VÉRIFICATION_SYSTÈMES_JOURNAL_RAPPORT_HEBDO + dateDebut.ToShortDateString() + "_Journal_Hebdo.csv";
            excelWB.SaveAs(nomJournal, Excel.XlFileFormat.xlCSV, misValue, misValue, misValue, misValue,
                           Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, true);

            excelWB.Close(true, nomJournal, misValue);

            Marshal.ReleaseComObject(excelWS);
            Marshal.ReleaseComObject(excelWB);
        }

        //***************************************************************************************************
        //* ConcatFichiers - Concaténation de tous les rapports journaliers de la semaine débutant par 
        //*                  dateDebut dans le fichier journal rapport hebdomadaire.
        //*
        //*                - dateDebut:  Date (DateTime) de début de la semaine désirée
        //*                - FichierToHebdo:   Liste (string) des fichiers à concatener
        //*                - FichierToHebdoTr: Liste (string) des fichiers à concatener
        //*
        //***************************************************************************************************
        private static void ConcatFichiers(DateTime dateDebut, List<string> FichierToHebdo, List<string> FichierToHebdoTr)
        {
            string nomJournal;

            nomJournal = PATH_VÉRIFICATION_SYSTÈMES_JOURNAL_RAPPORT_HEBDO + dateDebut.ToShortDateString() + "_Journal_Hebdo.csv";

            foreach (string F in FichierToHebdo)
            {
                File.AppendAllText(nomJournal, File.ReadAllText(PATH_VÉRIFICATION_SYSTÈMES_JOURNAL + F.Trim()));
            }
            //BGW_Loading_Indicator.ReportProgress(20);
            //Thread.Sleep(10000);
            foreach (string F in FichierToHebdoTr)
            {
                File.AppendAllText(nomJournal, File.ReadAllText(PATH_VÉRIFICATION_SYSTÈMES_JOURNAL_TRAITES + F.Trim()));
                //excelWBx[i] = excelApp.Workbooks.Open(PATH_VÉRIFICATION_SYSTÈMES_JOURNAL_TRAITES + F.Trim()); ET excelWS = excelWBx[1].Worksheets.get_Item(1);
                //i++;
            }
        }

        //***************************************************************************************************
        //* FormatFichier - Efface toutes les lignes contenant des entêtes et remaniement des colonnes
        //*
        //*               - dateDebut:  Date (DateTime) de début de la semaine désirée
        //*
        //***************************************************************************************************
        private static void FormatFichier(DateTime dateDebut)
        {
            string nomJournal;
            int lignesCompteur;
            int colonnesCompteur;
            Excel.Range currentFind;
            Excel.Range firstFind = null;
            Excel.Range signatures;
            SortedSet<int> lignesSupp = new SortedSet<int>();

            nomJournal = PATH_VÉRIFICATION_SYSTÈMES_JOURNAL_RAPPORT_HEBDO + dateDebut.ToShortDateString() + "_Journal_Hebdo.csv";

            // OUVERTURE EXCEL (Local:= True)
            excelWB = excelApp.Workbooks.Open(Filename: nomJournal, UpdateLinks: 2, ReadOnly: false, AddToMru: false, Notify: false, Local: true);
            excelWS = excelWB.ActiveSheet;

            xlRange = excelWS.UsedRange;

            //lignesCompteur = xlRange.Rows.Count; // var j = excelWS.UsedRange.Columns["A:A", Type.Missing].Rows.Count;
            //colonnesCompteur = xlRange.Columns.Count;

            lignesCompteur = excelWS.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            // Find the last real column
            colonnesCompteur = excelWS.Cells.Find("*", System.Reflection.Missing.Value,
                                           System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                           Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious,
                                           false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
            
            signatures = excelWS.Range[excelWS.Cells[2, 1], excelWS.Cells[lignesCompteur, colonnesCompteur]];
            /*currentFind = signatures.Find(excelWS.Cells[2, 1], System.Reflection.Missing.Value, Excel.XlFindLookIn.xlValues, System.Reflection.Missing.Value, 
                                          System.Reflection.Missing.Value, Excel.XlSearchDirection.xlNext, false, false, System.Reflection.Missing.Value); */

            
            /***** Effacer toutes les lignes d'entêtes *****/
            // You should specify all these parameters every time you call this method,
            // since they can be overridden in the user interface. 
            currentFind = signatures.Find(excelWS.Cells[2, 1], System.Reflection.Missing.Value, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                                          Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, System.Reflection.Missing.Value,
                                          System.Reflection.Missing.Value);

            lignesSupp.Add(currentFind.Row);

            while (currentFind != null)
            {
                // Keep track of the first range you find. 
                if (firstFind == null)
                {
                    firstFind = currentFind;
                }
                else
                { 
                    if (currentFind.get_Address(Excel.XlReferenceStyle.xlA1) == firstFind.get_Address(Excel.XlReferenceStyle.xlA1))
                    {
                        break;
                    }
                }
                currentFind.Font.Bold = true;
                currentFind = signatures.FindNext(currentFind);
                lignesSupp.Add(currentFind.Row);
                //BGW_Loading_Indicator.ReportProgress(30);
                //Thread.Sleep(1000);
            }

            foreach (int item in lignesSupp.Reverse())
            {

                signatures = excelWS.Range[excelWS.Cells[item, 1], excelWS.Cells[item, 1]];
                signatures.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            }
            /*********************************************/

            //Copier Colonne 2 dans colonne 3
            lignesCompteur = xlRange.Rows.Count;
            signatures = excelWS.Range[excelWS.Cells[2, 2], excelWS.Cells[lignesCompteur, 2]];
            signatures.Copy(excelWS.Cells[2, 3]);

            //Copier Colonne 1 dans colonne 2
            lignesCompteur = xlRange.Rows.Count;
            signatures = excelWS.Range[excelWS.Cells[2, 1], excelWS.Cells[lignesCompteur, 1]];
            signatures.Copy(excelWS.Cells[2, 2]);

            // Inscrire la date de début du cycle dans la première colonne
            excelWS.Cells[2, 1] = dateDebut.ToShortDateString();
            excelWS.Cells[2, 1].Copy(destination: excelWS.Range[excelWS.Cells[2, 1], excelWS.Cells[lignesCompteur, 1]]);

            //Écraser Colonne 4 avec Colonne 5
            signatures = excelWS.Range[excelWS.Cells[2, 5], excelWS.Cells[lignesCompteur, 5]];
            signatures.Copy(excelWS.Cells[2, 4]);

            //Écraser Colonne 6 avec Colonne 12
            signatures = excelWS.Range[excelWS.Cells[2, 12], excelWS.Cells[lignesCompteur, 12]];
            signatures.Copy(excelWS.Cells[2, 6]);

            //Éliminer les Colonnes 7 à 9 incl. et 11 à 13 incl.
            signatures = excelWS.Range[excelWS.Cells[1, 7], excelWS.Cells[lignesCompteur, 9]];// excelWS.Cells[lignesCompteur, 12]];
            signatures.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);
            signatures = excelWS.Range[excelWS.Cells[1, 8], excelWS.Cells[lignesCompteur, 10]];
            signatures.EntireColumn.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);
        }

        //***************************************************************************************************
        //* ClasserNoms - Trie ascendants selon l'index de la colonne désirée
        //*
        //*             - colIndex: Index de la colonne pour le tri
        //*
        //***************************************************************************************************
        private static void ClasserCol(ColId colIndex)
        {
            Excel.Range rangeNoms;
            Excel.Range rangeComplet;

            int lignesCompteur = xlRange.Rows.Count; // var j = excelWS.UsedRange.Columns["A:A", Type.Missing].Rows.Count;
            int colonnesCompteur = xlRange.Columns.Count;

            rangeComplet = excelWS.Range[excelWS.Cells[2, 1], excelWS.Cells[lignesCompteur, colonnesCompteur]];
            rangeNoms = excelWS.Range[excelWS.Cells[2, colIndex], excelWS.Cells[2, colIndex]];
            //excelWS.Range[rangeNoms].Select();
            rangeNoms = excelWS.Range[rangeNoms, rangeNoms.End[Excel.XlDirection.xlDown]];
            rangeNoms.Select();
            excelWS.Sort.SortFields.Clear();

            excelWS.Sort.SortFields.Add(rangeNoms, Excel.XlSortOn.xlSortOnValues, Excel.XlSortOrder.xlAscending, Excel.XlSortDataOption.xlSortNormal);
            {
                var withBlock = excelWS.Sort;
                withBlock.SetRange(rangeComplet);
                withBlock.Header = Excel.XlYesNoGuess.xlNo;
                withBlock.MatchCase = false;
                withBlock.Orientation = Excel.XlSortOrientation.xlSortColumns;
                withBlock.SortMethod = Excel.XlSortMethod.xlPinYin;
                withBlock.Apply();
            }
        }

        //***************************************************************************************************
        //* PivotMimic -  Sous le principe des PivotTable rassembler les violations par jeux, jours et 
        //*               finalement par catégories.
        //*
        //*
        //***************************************************************************************************
        private static void PivotMimic()
        {
            Excel.Range itemRecherche;
            SortedSet<int> resTrouves;
            SortedSet<string> nomsTrouves;
            SortedSet<DateTime> dateItem = new SortedSet<DateTime>();
            SortedSet<string> criticiteItem = new SortedSet<string>();

            //Classer par noms de la colonne 3
            ClasserCol(ColId.Nom);

            //Creer liste des noms de jeux
            itemRecherche = excelWS.Cells[2, ColId.Nom];
            nomsTrouves = CreeListeNomJeux(itemRecherche);

            foreach (var nom in nomsTrouves)
            {
                resTrouves = CreeListeItemsTrouves(nom);

                //BGW_Loading_Indicator.ReportProgress(60);
                //Thread.Sleep(100);

                if (resTrouves != null)
                {
                    for (int i = 0; i < (int)NbCat.NombreCAT; i++)
                    {
                        ligneCatxPrecedente[i] = 0; // CATx
                    }

                    dateItem.Clear();
                    criticiteItem.Clear();

                    foreach (int ligne in resTrouves.Reverse())
                    {
                        Excel.Range ligneDel;
                        excelWS.Cells[ligne, (int)ColId.Nb_de_jours_en_violations] = 0; //nb_Jours
                        
                        switch (excelWS.Cells[ligne, (int)ColId.Niveau_de_criticite].Value) // Niveau de criticité
                        {
                            case "CAT1":
                                ligneCatxPrecedente[(int)NbCat.CAT1] = AjusteNbJourCat(ligne, ligneCatxPrecedente[(int)NbCat.CAT1]);
                                break;

                            case "CAT2":
                                ligneCatxPrecedente[(int)NbCat.CAT2] = AjusteNbJourCat(ligne, ligneCatxPrecedente[(int)NbCat.CAT2]);
                                break;

                            case "CAT3":
                                ligneCatxPrecedente[(int)NbCat.CAT3] = AjusteNbJourCat(ligne, ligneCatxPrecedente[(int)NbCat.CAT3]);
                                break;

                            case "CAT4":
                                ligneCatxPrecedente[(int)NbCat.CAT4] = AjusteNbJourCat(ligne, ligneCatxPrecedente[(int)NbCat.CAT4]);
                                break;

                            default:
                                ligneDel = excelWS.Rows[ligne];
                                ligneDel.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                                for (int i = 0; i < (int)NbCat.NombreCAT; i++) // Nb de CAT
                                {
                                    if (ligneCatxPrecedente[i] > 0)
                                    {
                                        ligneCatxPrecedente[i]--;
                                    }
                                }
                                break;
                        }
                    }
                }
            }
        }

        //***************************************************************************************************
        //* AjusteNbJourCat - Traitement de ligne en cours (élimination de ligne avec addition si requise)
        //*                   pour ajustement du nombres de jours de violation.
        //*
        //*                 - ligneCatPrecedente: No de ligne précédente contenant la catégorie en cours
        //*                 - Retourne le nouveau no. de ligne en traitement
        //*
        //***************************************************************************************************
        private static int AjusteNbJourCat(int ligne, int ligneCatPrecedente)
        {
            Excel.Range ligneDel;

            if (ligneCatPrecedente == 0)
            {
                excelWS.Cells[ligne, (int)ColId.Nb_de_jours_en_violations] = excelWS.Cells[ligne, (int)ColId.Nb_de_jours_en_violations].Value + 1; //NB_Jours
            }
            else
            {
                if (excelWS.Cells[ligneCatPrecedente, (int)ColId.Date].Value == excelWS.Cells[ligne, (int)ColId.Date].Value) //Date
                {
                    excelWS.Cells[ligne, (int)ColId.Nb_de_jours_en_violations] = excelWS.Cells[ligneCatPrecedente, (int)ColId.Nb_de_jours_en_violations].Value; //NB_Jours
                }
                else
                {
                    excelWS.Cells[ligne, (int)ColId.Nb_de_jours_en_violations] = excelWS.Cells[ligneCatPrecedente, (int)ColId.Nb_de_jours_en_violations].Value + 1; //NB_Jours
                }
                ligneDel = excelWS.Rows[ligneCatPrecedente];
                ligneDel.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                /* Soustraire 1 de toutes les lignes précédentes pour chacunes des catégories */
                for (int i = 0; i < (int)NbCat.NombreCAT; i++) // Nb de CAT
                {
                    if (ligneCatxPrecedente[i] > ligneCatPrecedente && ligneCatxPrecedente[i] > 0)
                    {
                        ligneCatxPrecedente[i]--;
                    }
                }
            }
            return ligne;
        }

        //***************************************************************************************************
        //* CreeListeItemsTrouves - Création d'une liste triée en fonction d'un élément spécifié
        //*
        //*                       - itemRecherche: Élément spécifié pour la création de la liste
        //*                       - lignesItemTrouve: Retourne une liste (int) contenants les no. de lignes
        //*                         où ont été identifié l'élément spécifié
        //*
        //***************************************************************************************************
        private static SortedSet<int> CreeListeItemsTrouves(string itemRecherche)
        {
            int lignesCompteur;
            int colonnesCompteur;
            Excel.Range currentFind;
            Excel.Range firstFind = null;
            Excel.Range rangeToSearch;
            SortedSet<int> lignesItemTrouve = new SortedSet<int>();

            xlRange = excelWS.UsedRange;

            //lignesCompteur = xlRange.Rows.Count;
            //colonnesCompteur = xlRange.Columns.Count;

            lignesCompteur = excelWS.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            // Find the last real column
            colonnesCompteur = excelWS.Cells.Find("*", System.Reflection.Missing.Value,
                                           System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                           Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious,
                                           false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;

            rangeToSearch = excelWS.Range[excelWS.Cells[2, 3], excelWS.Cells[lignesCompteur, 3]]; // colonnesCompteur]];
            
            // You should specify all these parameters every time you call this method,
            // since they can be overridden in the user interface. 
            currentFind = rangeToSearch.Find(itemRecherche, System.Reflection.Missing.Value, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole,
                                          Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, System.Reflection.Missing.Value,
                                          System.Reflection.Missing.Value);

            //lignesItemTrouve.Add(currentFind.Row);

            while (currentFind != null)
            {
                lignesItemTrouve.Add(currentFind.Row);
                // Keep track of the first range you find. 
                if (firstFind == null)
                {
                    firstFind = currentFind;
                }
                else
                {
                    if (currentFind.get_Address(Excel.XlReferenceStyle.xlA1) == firstFind.get_Address(Excel.XlReferenceStyle.xlA1))
                    {
                        break;
                    }
                }
                currentFind.Font.Bold = true;
                currentFind = rangeToSearch.FindNext(currentFind);
                lignesItemTrouve.Add(currentFind.Row);
            }
            return lignesItemTrouve;
        }

        //***************************************************************************************************
        //* CreeListeNomJeux - Création d'une liste triée (sans duplication) contenant tous les noms de jeux
        //*
        //*                  - colonneNoms: Identifiant de la colonne du fichier Excel contenant les noms
        //*                                 de jeux pour la création de la liste
        //*                  - nomsItems: Retourne une liste (string) contenants tous les noms de jeux
        //*                               contenus dans le no. de colonne spécifiée
        //*
        //***************************************************************************************************
        private static SortedSet<string> CreeListeNomJeux(Excel.Range colonneNoms)
        {
            int colonneCourante, lignesCompteur;
            SortedSet<string> nomsItems = new SortedSet<string>();

            lignesCompteur = excelWS.Cells.Find("*", System.Reflection.Missing.Value,
                                                System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                                Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                                                false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            colonneCourante = colonneNoms.Cells.Column;

            for (int i = 2; i <= lignesCompteur; i++)
            {
                nomsItems.Add(excelWS.Cells[i, colonneCourante].Value());
            }

            return nomsItems;
        }

        //***************************************************************************************************
        //* CleanupJournal - Éliminer les informations inutiles du journal hebdo et appeler le tri par
        //*                  plateformes
        //*
        //***************************************************************************************************
        private static void CleanupJournal()
        {
            Excel.Range colonneDel;

            colonneDel = excelWS.Columns[(int)ColId.Evenement + 1];
            colonneDel.EntireColumn.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);
            colonneDel = excelWS.Columns[(int)ColId.Evenement];
            colonneDel.EntireColumn.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);
            colonneDel = excelWS.Columns[(int)ColId.Date];
            colonneDel.EntireColumn.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);

            // Trier par Plateformes
            ClasserCol(ColId.Plateforme);
    }

    //***************************************************************************************************
    //* Module_Verif_Excel - Termine les instances d'Excel existantes
    //*
    //*                
    //*
    //***************************************************************************************************
    private static void Module_Verif_Excel()
        {
            Debug.WriteLine("Module_Verif_Excel <ENTER>");

            try
            {
                Marshal.ReleaseComObject(excelWS);
                Marshal.ReleaseComObject(excelWB);
                Marshal.ReleaseComObject(excelApp);
            }
            catch (Exception)
            {
                excelWS = null;
                excelWB = null;
                excelApp = null;
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();

            try
            {
                Process[] obj1;// = new Process[10];
                obj1 = Process.GetProcessesByName("EXCEL");
                foreach (Process p in obj1)
                {
                    p.Kill();
                    System.Threading.Thread.Sleep(100);
                }
            }
            catch (Exception)
            {
                Debug.WriteLine("Plus d'Excel actif");
            }

            Debug.WriteLine("Module_Verif_Excel <EXIT>");
        }

        //***************************************************************************************************
        //* FermetureExcel - Sauve et ferme le journal; s'assure que toutes les instances d'Excel soient 
        //*                  terminées
        //*
        //***************************************************************************************************
        private static void FermetureExcel(string nomJournal)
        {
            //string nomJournal;
            object misValue = System.Reflection.Missing.Value;

            //nomJournal = PATH_VÉRIFICATION_SYSTÈMES_JOURNAL_RAPPORT_HEBDO + dateDebut.ToShortDateString() + "_Journal_Hebdo.csv";
            excelWB.SaveAs(nomJournal, Excel.XlFileFormat.xlCSV, misValue, misValue, misValue, misValue,
                           Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, true);
#if DEBUG
            MessageBox.Show("Fichier Excel créé: " + nomJournal);
#endif
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(excelWS);

            //close and release
            excelWB.Close();
            Marshal.ReleaseComObject(excelWB);

            //quit and release
            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);

            Module_Verif_Excel();
        }
    }
}
