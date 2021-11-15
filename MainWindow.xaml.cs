using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Diagnostics;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace HebdoJEL
{
    /// <summary>
    /// Logique d'interaction pour MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //static int i = 0;
        DateTime date;
        readonly DateManqModel dateManqModel = new DateManqModel();

        public MainWindow()
        {
            InitializeComponent();

            //this.Title = "Bilan Hebdomadaire de Jeux";
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ShowMissingDates.DataContext = dateManqModel;
        }

        //public string Texte { get; set; } = "Hello!";

        private void Depart(object sender, RoutedEventArgs e)
        {
            Date_Debut.DisplayDateEnd = DateTime.Today.AddDays(-7);
            Bouton_Generer.IsEnabled = false;
            dateManqModel.DateManq = "";
        }

        //***************************************************************************************************
        //* Calendar_SelectedDatesChanged - Sélectionne les dates formant la semaine débutant par la date 
        //*                                 sélectionnée
        //*
        //***************************************************************************************************
        private void Calendar_SelectedDatesChanged(object sender, SelectionChangedEventArgs e)
        {
            List<string> fichierManquant;
            string jourManq;
            int JJ, MM, YY;
            DateTime jourSansJournal;
            Debug.WriteLine("Calendar_SelectedDatesChanged <ENTER>");

            date = Date_Debut.SelectedDate.Value;
            Date_Debut.BlackoutDates.Clear();
            fichierManquant = App.DateValidation(date);
            
            if (fichierManquant.Count != 0)
            {
                dateManqModel.DateManq = "Date(s) Manquante(s) = ";
                foreach (var fM in fichierManquant)
                {
                    jourSansJournal = DateTime.Parse(fM.Substring(0, 10));
                    jourManq = jourSansJournal.ToShortDateString();
                    YY = int.Parse(jourManq.Substring(0, 4));
                    MM = int.Parse(jourManq.Substring(5, 2));
                    JJ = int.Parse(jourManq.Substring(8, 2));
                    if (Date_Debut.SelectedDate.Value != jourSansJournal)
                    {
                        Date_Debut.BlackoutDates.Add(new CalendarDateRange(new DateTime(YY, MM, JJ)));
                    }
                    dateManqModel.DateManq += fM + " ,";
                    ShowMissingDates.Foreground = new SolidColorBrush(Colors.Red);
                }
                Bouton_Generer.IsEnabled = false;
                MessageBox.Show("Semaine incomplète!", "HebdoJEL", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else
            {
                Bouton_Generer.IsEnabled = true;
                dateManqModel.DateManq = "";
                MessageBox.Show("Cliquer sur <Générer> pour la création du rapport de la semaine du " + date.ToString("d") + " sinon choisir une autre date.");
            }
            
            Debug.WriteLine("Calendar_SelectedDatesChanged <EXIT>");
        }

        //***************************************************************************************************
        //* Bouton_Generer_Click - Démarre la journalisation avec les dates
        //*                        sélectionnées
        //*
        //***************************************************************************************************
        private void Bouton_Generer_Click(object sender, RoutedEventArgs e)
        {
            ShowMissingDates.FontSize = 11;
            ShowMissingDates.Foreground = new SolidColorBrush(Colors.OrangeRed);
            dateManqModel.DateManq = "Rapport du " + date.ToString("d") + " en création";
            MessageBox.Show("La génération du rapport peu prendre un peu de temps, merci de patienter!");
            bool v = App.JournalHebdo(date);
            Bouton_Generer.IsEnabled = false;
            dateManqModel.DateManq = "";
            if (v)
            {
                MessageBox.Show("Journal Hebdomadaire semaine du " + date.ToString("d") + " créé.");
            }
            Bouton_Generer.IsEnabled = false;
        }
    }

    // Classes
    public class DateManqModel : INotifyPropertyChanged
    {
        private string _dateManq;

        public string DateManq
        {
            get { return _dateManq; }
            set
            {
                _dateManq = value;
                OnPropertyChanged();
                //OnPropertyChanged(nameof(DateManq));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        //protected void OnPropertyChanged(string propertyName)
        //{
        //    if (PropertyChanged != null)
        //    {
        //        PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        //    }
        //}
    }
}
