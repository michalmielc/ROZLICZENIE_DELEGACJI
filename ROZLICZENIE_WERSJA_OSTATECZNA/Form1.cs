using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.IO;
using System.Xml;
using System.Web;
using System.Net;
using System.Reflection;
using iTextSharp.text.pdf;
using itextsharp.pdfa;
using iTextSharp.text;
using iTextSharp;

namespace ROZLICZENIE_WERSJA_OSTATECZNA
{
    public partial class Form1 : Form
    {
        public string wersja = "wersja 2.4";   //ZMIENNA GLOBALNA DRUKUJE SIĘ NA KAŻDYM FORMULARZU


        public bool internetConncetion = true;

        public Form1()
        {
            InitializeComponent();
         
    }


    // ŁADOWANIE PROGRAMU ##################################################################


    // zapis danych nagłówka
    private void zmianyZmiany_Click1(object sender, EventArgs e)
        {

            progressBar1.Visible = true;
            XmlDocument doc;
            XmlElement root;
            doc = new XmlDocument();
            doc.Load((@System.IO.Directory.GetCurrentDirectory().ToString() + "\\ustawienia.xml"));
            root = doc.DocumentElement;

            root.GetElementsByTagName("name")[0].InnerText = textBox1.Text;
            root.GetElementsByTagName("idp")[0].InnerText = textBox2.Text;
            root.GetElementsByTagName("mpk")[0].InnerText = textBox3.Text;
            root.GetElementsByTagName("nrkonta")[0].InnerText = textBox4.Text;
            root.GetElementsByTagName("okres")[0].InnerText = textBox5.Text;
            root.GetElementsByTagName("visa")[0].InnerText = textBox6.Text;
            doc.Save((@System.IO.Directory.GetCurrentDirectory().ToString() + "\\ustawienia.xml"));


            progressBar1.Value = 1;

            for (int i = 0; i < 100; i++)
            {

                Thread.Sleep(1);
                progressBar1.Increment(15);
            }

            progressBar1.Visible = false;

        }

        // wczytanie danych nagłówka

        void wczytaj_daneXML()
        {

            XmlDocument doc;
            XmlElement root;
            doc = new XmlDocument();
            doc.Load((@System.IO.Directory.GetCurrentDirectory().ToString() + "\\ustawienia.xml"));
            root = doc.DocumentElement;

            textBox1.Text = root.GetElementsByTagName("name")[0].InnerText;
            textBox2.Text = root.GetElementsByTagName("idp")[0].InnerText;
            textBox3.Text = root.GetElementsByTagName("mpk")[0].InnerText;
            textBox4.Text = root.GetElementsByTagName("nrkonta")[0].InnerText;
            textBox5.Text = root.GetElementsByTagName("okres")[0].InnerText;
            textBox6.Text = root.GetElementsByTagName("visa")[0].InnerText;

        }


        // ładowanie formularza - ukrycie wszystkich tabel
        private void Form1_Load(object sender, EventArgs e)
        {
            this.Text = "FORMULARZ ROZLICZEŃ " + wersja;

            wczytaj_daneXML();
            tabeleKrajowa.Visible = false;
            tabeleZagraniczna.Visible = false;
            mojeDane.Visible = false;
        

            wczytaj_tabeleXML(noclegiKrajowa);
            wczytaj_tabeleXML(dodatkoweKrajowa);
            wczytaj_tabeleXML(samochoduKrajowa);
            wczytaj_tabeleXML(delegacjeKrajowa);
     

            //wczytanie kont dla dodZagr
            wczytajkontadodZagr();

            //wczytanie kont dla samZagr
            wczytajkontasamZagr();

            //wczytanie walut do Comboboxów
            podstawowe_waluty();

            //ladowanie kursow

            loadChangerate();
    
            //wczytanie stawek zagranicznych

            wczytajstawkiZagr();

            //wczytanie sum tabel
            label17.Text = sumuj(noclegiKrajowa, 2).ToString();
            label18.Text = sumuj(noclegiKrajowa, 3).ToString("N2");
            label20.Text = sumuj(dodatkoweKrajowa, 2).ToString("N2");
            label22.Text = sumuj(samochoduKrajowa, 2).ToString("N2");
            label35.Text = sumuj(delegacjeKrajowa, 10).ToString("N2");
            label27.Text = sumuj(noclegiZagraniczna, 6).ToString("N2");
            label25.Text = sumuj(noclegiZagraniczna, 2).ToString();
            label44.Text = sumuj(dodatkoweZagraniczna, 5).ToString("N2");
            label1.Text = sumuj(samochoduZagraniczna, 5).ToString("N2");
            label52.Text = sumuj(delegacjeZagraniczna, 13).ToString("N2");


        }

        //laduj kursy
        void loadChangerate()
        {
            try
            {
                System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                System.Net.WebClient wb = new System.Net.WebClient();
                wb.DownloadFile("http://www.nbp.pl/Kursy/xml/dir.txt", "dir.txt");
                read_currency("LastA");
                internetConncetion = true;

            }

            catch (Exception ex)
            {
                MessageBox.Show("Brak połączenia z internetem lub strona NBP nie odpowiada. Rozliczenie delegacji zagranicznej nie będzie możliwe!");
                MessageBox.Show(ex.Message + ex.ToString() + ex.Source);


                internetConncetion = false;
                tabeleZagraniczna.Visible = false;
            }

        }

        // ładowanie podstawowych walut//

        void podstawowe_waluty()
        {

            //waluty
            string[] waluty = new string[] { "?","EUR", "PLN","CHF", "USD", "HUF", "UAH", "CZK", "HRK", "NOK", "SEK", "DKK", "RON", "RUB", "BGN","GBP" };

            comboBox6.Items.AddRange(waluty);
            comboBox7.Items.AddRange(waluty);
            comboBox10.Items.AddRange(waluty);
      
        }
           

        void wczytajkontadodZagr()
        {

            comboBox8.Items.Add("-- WYBIERZ KOSZT --");         
            comboBox8.Items.Add("Artykuły biurowe za granicą");
            comboBox8.Items.Add("Rozmowy handlowe za granicą - bowling");
            comboBox8.Items.Add("Rozmowy handlowe za granicą - rachunki z restauracji");
            comboBox8.Items.Add("Opłata parkingowa za granicą");
            comboBox8.Items.Add("Pozostałe koszty dodatkowe podróży za granicą");
            comboBox8.Items.Add("Koszty taksówek za granicą");
            comboBox8.Items.Add("Bilet wstępu na targi za granicą");
            comboBox8.Items.Add("Prom, opłata drogowa, pociąg, samolot, autostrada, vinieta za granicą");
            comboBox8.Items.Add("Wynajem samochodu za granicą");
            comboBox8.Items.Add("Napiwki za granicą");
            comboBox8.Items.Add("Poczęstunek. Rozmowy handlowe na terenie kraju");
            comboBox8.Items.Add("Oświadczenie");
            comboBox8.Items.Add("Wydatki za zakup wyposażenia");

        }


        void wczytajkontasamZagr()
     {

  
         comboBox9.Items.Add("-- WYBIERZ KOSZT --");
         comboBox9.Items.Add("Koszty eksploatacji sam. os. (paliwo + akcesoria) za granicą");
         comboBox9.Items.Add("Naprawy/Myjnia do sam. za granicą");
         comboBox9.Items.Add("Inne usługi samochodowe za granicą");
        }

        void wczytajstawkiZagr()
     {
         comboBox11.Items.Add("-- WYBIERZ KRAJ --");
         comboBox11.Items.Add("NIEMCY");
         comboBox11.Items.Add("AUSTRIA");
         comboBox11.Items.Add("BELGIA");
         comboBox11.Items.Add("BUŁGARIA");
         comboBox11.Items.Add("BOŚNIA I HERCEGOWINA");
         comboBox11.Items.Add("BIAŁORUŚ");
         comboBox11.Items.Add("SZWAJCARIA");
         comboBox11.Items.Add("CZECHY");
         comboBox11.Items.Add("WĘGRY");
         comboBox11.Items.Add("FRANCJA");
         comboBox11.Items.Add("WIELKA BRYTANIA");
         comboBox11.Items.Add("LUKSEMBURG");
         comboBox11.Items.Add("LITWA");
         comboBox11.Items.Add("NORWEGIA");
         comboBox11.Items.Add("HOLANDIA");
         comboBox11.Items.Add("RUMUNIA");
         comboBox11.Items.Add("SZWECJA");
         comboBox11.Items.Add("SŁOWACJA");
         comboBox11.Items.Add("UKRAINA");
         comboBox11.Items.Add("ŁOTWA");
         comboBox11.Items.Add("ESTONIA");
         comboBox11.Items.Add("DANIA");

         comboBox5.Items.Add("-- WYBIERZ KRAJ --");
         comboBox5.Items.Add("NIEMCY");
         comboBox5.Items.Add("AUSTRIA");
         comboBox5.Items.Add("BELGIA");
         comboBox5.Items.Add("BUŁGARIA");
         comboBox5.Items.Add("BOŚNIA I HERCEGOWINA");
         comboBox5.Items.Add("BIAŁORUŚ");
         comboBox5.Items.Add("SZWAJCARIA");
         comboBox5.Items.Add("CZECHY");
         comboBox5.Items.Add("WĘGRY");
         comboBox5.Items.Add("FRANCJA");
         comboBox5.Items.Add("WIELKA BRYTANIA");
         comboBox5.Items.Add("LUKSEMBURG");
         comboBox5.Items.Add("LITWA");
         comboBox5.Items.Add("NORWEGIA");
         comboBox5.Items.Add("HOLANDIA");
         comboBox5.Items.Add("RUMUNIA");
         comboBox5.Items.Add("SZWECJA");
         comboBox5.Items.Add("SŁOWACJA");
         comboBox5.Items.Add("UKRAINA");
         comboBox5.Items.Add("ŁOTWA");
         comboBox5.Items.Add("ESTONIA");
         comboBox5.Items.Add("DANIA");




  stawkiZagraniczna.Rows.Add("NIEMCY","EUR","16,33","24,50","49,00","37,50");
  stawkiZagraniczna.Rows.Add("AUSTRIA", "EUR", "19,00", "28,50", "57,00", "32,50");
  stawkiZagraniczna.Rows.Add("BELGIA", "EUR", "16,00", "24,00", "48,00", "40,00");
  stawkiZagraniczna.Rows.Add("BUŁGARIA", "EUR", "14,33", "21,50", "43,00", "30,00");
  stawkiZagraniczna.Rows.Add("BOŚNIA I HERCEGOWINA", "EUR", "13,67", "20,50", "41,00", "25,00");
  stawkiZagraniczna.Rows.Add("BIAŁORUŚ","EUR","14,00","21,00","42,00","52,00");
  stawkiZagraniczna.Rows.Add("SZWAJCARIA","CHF","29,33","44,00","88,00","50,00");
  stawkiZagraniczna.Rows.Add("CZECHY", "EUR", "13,67", "20,50", "41,00", "30,00");
  stawkiZagraniczna.Rows.Add("WĘGRY", "EUR", "14,67", "22,00", "44,00", "32,50");
  stawkiZagraniczna.Rows.Add("FRANCJA", "EUR", "18,33", "27,50", "55,00", "45,00");
  stawkiZagraniczna.Rows.Add("WIELKA BRYTANIA", "GBP", "11,67", "17,50", "35,00", "50,00");
  stawkiZagraniczna.Rows.Add("CHORWACJA", "EUR", "14,00", "21,00", "42,00", "31,25");
  stawkiZagraniczna.Rows.Add("WŁOCHY", "EUR", "16,00", "24,00", "48,00", "43,50");
  stawkiZagraniczna.Rows.Add("LUKSEMBURG", "EUR", "16,00", "24,00", "48,00", "40,00");
  stawkiZagraniczna.Rows.Add("LITWA", "EUR", "15,00", "22,50", "45,00", "32,50");
  stawkiZagraniczna.Rows.Add("NORWEGIA", "NOK", "150,32", "225,50", "451,00", "375,00");
  stawkiZagraniczna.Rows.Add("HOLANDIA", "EUR", "16,67", "25,00", "50,00", "32,50");
  stawkiZagraniczna.Rows.Add("RUMUNIA", "EUR	", "12,67", "19,00", "38,00", "25,00");
  stawkiZagraniczna.Rows.Add("SZWECJA", "SEK", "170,00", "255,00", "510,00", "450,00");
  stawkiZagraniczna.Rows.Add("SŁOWACJA", "EUR", "15,67", "23,50", "47,00", "30,00");
  stawkiZagraniczna.Rows.Add("UKRAINA", "EUR", "13,67", "20,50", "41,00", "45,00");
  stawkiZagraniczna.Rows.Add("ŁOTWA", "EUR", "19,00", "28,50", "57,00", "33,00");
  stawkiZagraniczna.Rows.Add("ESTONIA", "EUR", "13,67", "20,50", "41,00", "45,00");
  stawkiZagraniczna.Rows.Add("DANIA", "DKK", "148,67", "223,00", "446,00", "1300,00");




        }
        
        
        //  - KONIEC  ŁADOWANIA PROGRAMU ##################################################################

        // MENU - ##################################################################


        //AKTYWACJA TABEL KRAJOWA
        private void dELEGACJAKRAJOWAToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabeleKrajowa.Visible = true;
            tabeleKrajowa.Location = new Point(12, 42);
            tabeleKrajowa.Size = new Size(1300, 600);
         
            tabeleZagraniczna.Visible = false;
            mojeDane.Visible = false;
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = 0;
            comboBox4.SelectedIndex = 0;
            comboBox20.SelectedIndex = 0;

            readLastfile.Visible = true;
            zapisdoXML.Visible = true;
            
            
            label7.Text = " ROZLICZENIE ZALICZKI/ DELEGACJI KRAJOWEJ ZA OKRES:  " + textBox5.Text;

        }

        //AKTYWACJA TABEL ZAGRANICZNA
        private void rOZLICZENIEZALICZKIDELEGACJIZAGRANICZNAToolStripMenuItem_Click(object sender, EventArgs e)
        {

            tabeleKrajowa.Visible = false;
            tabeleZagraniczna.Visible = false;

            loadChangerate();

            if(!internetConncetion)
            {
                MessageBox.Show("BRAK POŁĄCZENIA Z INTERNETEM. ROZLICZENIE DEL. ZGARANICZNEJ NIE JEST MOZLIWE!");
                
                return;
            }

          
            

            if (Form2.pierwszeWywolanie == true)
            {
                DialogResult result = MessageBox.Show("Zmiana kursu spowoduje usunięcie wpisów w tabeli!!!", "ZAPYTANIE", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

                if (result == DialogResult.Cancel)
                {
                    return;
                }

                else
                {

                    czysctabele();
                }

            }


         
                read_currency("LastA");
          
            
            if (internetConncetion)
                {
                Form2 form2 = new Form2();
                form2.ShowDialog(this);
                internetConncetion = true;
            }

            else
            {
                return;
            }


            read_currency("LastA");
            if (internetConncetion)
            {
                if (Form2.pierwszeWywolanie == true)
                {


                    read_currency(Form2.nazwaTabeli);


                    tabeleZagraniczna.Visible = true;
                    tabeleZagraniczna.Location = new Point(12, 42);
                    tabeleZagraniczna.Size = new Size(1400, 600);
                    mojeDane.Visible = false;
                    label7.Text = "ROZLICZENIE ZALICZKI/ DELEGACJI ZAGRANICZNEJ ZA OKRES:  " + textBox5.Text;
                    comboBox6.SelectedIndex = 0;
                    comboBox7.SelectedIndex = 0;
                    comboBox8.SelectedIndex = 0;
                    comboBox9.SelectedIndex = 0;
                    comboBox10.SelectedIndex = 0;
                    comboBox11.SelectedIndex = 1;
                    comboBox13.SelectedIndex = 0;
                    comboBox14.SelectedIndex = 0;
                    comboBox15.SelectedIndex = 0;
                    comboBox16.SelectedIndex = 0;
                    comboBox12.SelectedIndex = 0;
                    comboBox17.SelectedIndex = 0;
                    comboBox18.SelectedIndex = 0;
                    comboBox19.SelectedIndex = 0;

                    zapisdoXML.Visible = false;
                    readLastfile.Visible = false;

                }
            
            }

                else
                {
                    internetConncetion = false;
                    return;
                }
            
        }

        //AKTYWACJA MOICH DANYCH
        private void mOJEDANEToolStripMenuItem_Click(object sender, EventArgs e)
        {

            mojeDane.Visible = true;
            mojeDane.Location = new Point(13, 27);
            mojeDane.Size = new Size(750, 400);
            tabeleKrajowa.Visible = false;
            tabeleZagraniczna.Visible = false;
            label7.Text = "DELEGACJA KRAJOWA ZA OKRES:  " + textBox5.Text;
        }


        // KONIEC MENU - ##########################################################################################


        // DELEGACJA KRAJOWA   ZMIANY COMBOBOX-A ##################################################################

        // TABELA NOCLEGI

        private void noclegiKrajowa_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            ComboBox cb1 = e.Control as ComboBox;

            if (noclegiKrajowa.CurrentCell.ColumnIndex == 1)
            {
                cb1.SelectedIndexChanged -= new
                 EventHandler(cb1_SelectedIndexChanged);

                cb1.SelectedIndexChanged += new
                 EventHandler(cb1_SelectedIndexChanged);

            }
        }


        void cb1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int selldx = ((ComboBox)sender).SelectedIndex;

    

            if (selldx == 0)
            {
                noclegiKrajowa.Rows[noclegiKrajowa.CurrentRow.Index].Cells[4].Value = 66660;

            }

            else
            {
                noclegiKrajowa.Rows[noclegiKrajowa.CurrentRow.Index].Cells[4].Value = 66640;
                noclegiKrajowa.Rows[noclegiKrajowa.CurrentRow.Index].Cells[3].Value = 67.5;
                noclegiKrajowa.Rows[noclegiKrajowa.CurrentRow.Index].Cells[2].Value = 1;
            }


        }

        // TABELA KOSZTY DODATKOWE

        private void dodatkoweKrajowa_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {

            ComboBox cb2 = e.Control as ComboBox;

            if (dodatkoweKrajowa.CurrentCell.ColumnIndex == 1)
            {

                // first remove event handler to keep from attaching multiple:


                cb2.SelectedIndexChanged -= new

                EventHandler(cb2_SelectedIndexChanged);

                cb2.SelectedIndexChanged += new

             EventHandler(cb2_SelectedIndexChanged);

            }
        }

        void cb2_SelectedIndexChanged(object sender, EventArgs e)
        {

            int selldx = ((ComboBox)sender).SelectedIndex;


            switch (selldx)
            {

                case 0:
                case 1:
                case 2:
                case 3:
                    dodatkoweKrajowa.Rows[dodatkoweKrajowa.CurrentRow.Index].Cells[3].Value = 68150;

                    break;
                case 4:
                case 5:
                case 6:
                case 7:
                case 8:
                case 9:

                    dodatkoweKrajowa.Rows[dodatkoweKrajowa.CurrentRow.Index].Cells[3].Value = 63000;

                    break;
                case 10:
                case 11:
                case 12:
                case 13:
                case 14:
                case 15:
                case 16:
                case 17:
                case 18:

                    dodatkoweKrajowa.Rows[dodatkoweKrajowa.CurrentRow.Index].Cells[3].Value = 66630;

                    break;

                case 19:

                    dodatkoweKrajowa.Rows[dodatkoweKrajowa.CurrentRow.Index].Cells[3].Value = 68000;

                    break;

                case 20:

                    dodatkoweKrajowa.Rows[dodatkoweKrajowa.CurrentRow.Index].Cells[3].Value = 66400;

                    break;

                case 21:
                case 22:

                    dodatkoweKrajowa.Rows[dodatkoweKrajowa.CurrentRow.Index].Cells[3].Value = 66408;

                    break;

                case 23:

                    dodatkoweKrajowa.Rows[dodatkoweKrajowa.CurrentRow.Index].Cells[3].Value = 65500;

                    break;
                case 24:

                    dodatkoweKrajowa.Rows[dodatkoweKrajowa.CurrentRow.Index].Cells[3].Value = 68050;

                    break;

                case 25:

                    dodatkoweKrajowa.Rows[dodatkoweKrajowa.CurrentRow.Index].Cells[3].Value = 68200;

                    break;
                case 26:

                    dodatkoweKrajowa.Rows[dodatkoweKrajowa.CurrentRow.Index].Cells[3].Value = 66120;

                    break;
                case 27:

                    dodatkoweKrajowa.Rows[dodatkoweKrajowa.CurrentRow.Index].Cells[3].Value = 66108;

                    break;
                case 28:
                case 31:
                case 32:
                    dodatkoweKrajowa.Rows[dodatkoweKrajowa.CurrentRow.Index].Cells[3].Value = 68210;

                    break;
                case 29:
                case 30:
                    dodatkoweKrajowa.Rows[dodatkoweKrajowa.CurrentRow.Index].Cells[3].Value = 66420;

                    break;

                case 33:

                    dodatkoweKrajowa.Rows[dodatkoweKrajowa.CurrentRow.Index].Cells[3].Value = 66660;
                    break;
                case 34:

                    dodatkoweKrajowa.Rows[dodatkoweKrajowa.CurrentRow.Index].Cells[3].Value = 68160;
                    break;

                default:
                    break;


            }

        }

        // TABELA KOSZTY SAMOCHODU

        private void samochoduKrajowa_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {

            ComboBox cb3 = e.Control as ComboBox;

            if (samochoduKrajowa.CurrentCell.ColumnIndex == 1)
            {

                // first remove event handler to keep from attaching multiple:



                cb3.SelectedIndexChanged -= new

                EventHandler(cb3_SelectedIndexChanged);

                cb3.SelectedIndexChanged += new

             EventHandler(cb3_SelectedIndexChanged);

            }
        }

        void cb3_SelectedIndexChanged(object sender, EventArgs e)
        {


            int selldx = ((ComboBox)sender).SelectedIndex;


          

            if (selldx == 0)
            {
                samochoduKrajowa.Rows[samochoduKrajowa.CurrentRow.Index].Cells[3].Value = 65100;
            }

            else
            {
                samochoduKrajowa.Rows[samochoduKrajowa.CurrentRow.Index].Cells[3].Value = 65000;
            }


        }


        // KONIEC  - DELEGACJA KRAJOWA   ZMIANY COMBOBOX-A ##################################################################



   
        string szukajKursu( string waluta)
        {
            string kurs = "";
            double przelicznik=1;
            double kurswaluty=0;

            if (waluta == "PLN")
            
            {
                return przelicznik.ToString("N4");

            }
            
            foreach (DataGridViewRow row in kursy.Rows)
            {
                if (row.Cells[2].Value.ToString().Equals(waluta))
                {
                    kurswaluty = double.Parse(row.Cells[3].Value.ToString());
                    przelicznik = double.Parse(row.Cells[1].Value.ToString());
                   
                    kurs = (kurswaluty/przelicznik).ToString("N4");
                    break;
                }
            }

            
            return kurs;


        }



        void szukajRyczaltu(string kraj)
        {
            string ryczalt = "";
     
            
            foreach (DataGridViewRow row in stawkiZagraniczna.Rows)
            {
                if (row.Cells[0].Value.ToString().Equals(kraj))
                {

         
                 comboBox6.Text= row.Cells[1].Value.ToString();
                 textBox9.Text = szukajKursu(comboBox6.Text);

                 ryczalt = double.Parse(row.Cells[5].Value.ToString()).ToString("N2");
                    break;
                }
            }



            textBox8.Text= ryczalt;


        }


        //  KONIEC  - TABELE ZAGRANICZNE - COMBOBOX ####################################################################################

        // //

        //

        //  TABELE KRAJOWE I ZAGRANICZNE USUŃ WIERSZ ####################################################################################

        public void usun_wiersz(DataGridView dgv)
        {
            if (dgv.Rows.Count > 1)
            {

                if (dgv.CurrentCell.ColumnIndex == 0 && dgv.CurrentCell.RowIndex < dgv.Rows.Count - 1)
                {


                    dgv.Rows.RemoveAt(dgv.CurrentCell.RowIndex);


                }

            }
        }


        private void noclegiKrajowa_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            usun_wiersz(noclegiKrajowa);
            label17.Text = sumuj(noclegiKrajowa, 2).ToString();
            label18.Text = sumuj(noclegiKrajowa, 3).ToString("N2");
        }

        private void dodatkoweKrajowa_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            usun_wiersz(dodatkoweKrajowa);

            label20.Text = sumuj(dodatkoweKrajowa, 2).ToString("N2");
        }

        private void samochoduKrajowa_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            usun_wiersz(samochoduKrajowa);
            label22.Text = sumuj(samochoduKrajowa, 2).ToString("N2");
        }

        private void delegacjeKrajowa_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            usun_wiersz(delegacjeKrajowa);
            label35.Text = sumuj(delegacjeKrajowa, 10).ToString("N2");
        }


        private void noclegiZagraniczna_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            usun_wiersz(noclegiZagraniczna);
            label27.Text = sumuj(noclegiZagraniczna, 6).ToString("N2");
            label25.Text = sumuj(noclegiZagraniczna, 2).ToString();
        
        }

        private void dodatkoweZagraniczna_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            usun_wiersz(dodatkoweZagraniczna);

            label44.Text = sumuj(dodatkoweZagraniczna, 5).ToString("N2");
        }

        private void samochoduZagraniczna_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            usun_wiersz(samochoduZagraniczna);
            label1.Text = sumuj(samochoduZagraniczna, 5).ToString("N2");
        }

        private void delegacjeZagraniczna_CellClick(object sender, DataGridViewCellEventArgs e)
        {


            usun_wiersz(delegacjeZagraniczna);
            label52.Text = sumuj(delegacjeZagraniczna, 14).ToString("N2");
        }

        //  KONIEC  - TABELE KRAJOWE ZAGRANICZNE -USUŃ WIERSZ####################################################################################


    
        // TABELE KRAJOWE - edycja pól // ################################################

        private void noclegiKrajowa_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (noclegiKrajowa.CurrentRow.Index != noclegiKrajowa.Rows.Count - 1)
            {
                formatowanieDwamiejsca(noclegiKrajowa, 3);
                formatowanieInt(noclegiKrajowa, 2);
            }


            if (noclegiKrajowa.CurrentRow.Cells[1].Value !=null)
            {
                if (noclegiKrajowa.CurrentRow.Cells[1].Value.ToString() == "Ryczałt za nocleg")
                {


                    formatowanieDwamiejsca(noclegiKrajowa, 3);
                    formatowanieInt(noclegiKrajowa, 2);
                    int il = int.Parse(noclegiKrajowa.Rows[noclegiKrajowa.CurrentRow.Index].Cells[2].Value.ToString());
                    noclegiKrajowa.Rows[noclegiKrajowa.CurrentRow.Index].Cells[3].Value = (67.5 * il).ToString("N2");

                }

                }

            
   
            

                label17.Text = sumuj(noclegiKrajowa, 2).ToString();
                label18.Text = sumuj(noclegiKrajowa, 3).ToString("N2");

            
        }

        private void dodatkoweKrajowa_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (dodatkoweKrajowa.CurrentRow.Index != dodatkoweKrajowa.Rows.Count - 1)
            {
                formatowanieDwamiejsca(dodatkoweKrajowa, 2);
            }

            label20.Text = sumuj(dodatkoweKrajowa, 2).ToString("N2");
        }

        private void samochoduKrajowa_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
              if (samochoduKrajowa.CurrentRow.Index != samochoduKrajowa.Rows.Count - 1)
            {
                formatowanieDwamiejsca(samochoduKrajowa, 2);
            }

            label22.Text = sumuj(samochoduKrajowa, 2).ToString("N2");
        }

        private void noclegiZagraniczna_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (noclegiZagraniczna.CurrentRow.Index != noclegiZagraniczna.Rows.Count - 1)
            {

                formatowanieInt(noclegiZagraniczna, 2);
                formatowanieDwamiejsca(noclegiZagraniczna, 3);
                formatowanieDwamiejsca(noclegiZagraniczna, 6);
            }

            label25.Text = sumuj(noclegiZagraniczna, 2).ToString();
            label27.Text = sumuj(noclegiZagraniczna, 6).ToString();
        }


        // KONIEC- edycja pól // ################################################





        //  - - - funkcje techniczne ----------------------------------------------------------------
        private void czyscTabele_Click(object sender, EventArgs e)
        {

            czysctabele();


        }


        void czysctabele()
        {

            if (noclegiKrajowa.Visible == true)
            {
                noclegiKrajowa.Rows.Clear();
                label17.Text = sumuj(noclegiKrajowa, 2).ToString();
                label18.Text = sumuj(noclegiKrajowa, 3).ToString("N2");

            }

            if (dodatkoweKrajowa.Visible == true)
            {
                dodatkoweKrajowa.Rows.Clear();
                label20.Text = sumuj(dodatkoweKrajowa, 2).ToString("N2");

            }

            if (samochoduKrajowa.Visible == true)
            {
                samochoduKrajowa.Rows.Clear();
                label22.Text = sumuj(samochoduKrajowa, 2).ToString("N2");
            }

            if (delegacjeKrajowa.Visible == true)
            {
                delegacjeKrajowa.Rows.Clear();
                label35.Text = sumuj(delegacjeKrajowa, 10).ToString("N2");
            }



            if (noclegiZagraniczna.Visible == true)
            {
                noclegiZagraniczna.Rows.Clear();
                label25.Text = "0";
                label27.Text = "0,00";


            }

            if (dodatkoweZagraniczna.Visible == true)
            {
                dodatkoweZagraniczna.Rows.Clear();

            }

            if (samochoduZagraniczna.Visible == true)
            {
                samochoduZagraniczna.Rows.Clear();

       
            }

            if (delegacjeZagraniczna.Visible == true)
            {
                delegacjeZagraniczna.Rows.Clear();

            }

        }

        // formatowanie dwa miejsca po przecinku

        void formatowanieDwamiejsca(DataGridView dgv, int kolumna)
        {

            if (dgv.Rows.Count == 1)
            {
                return;
            }

            if (dgv.CurrentRow.Cells[kolumna].Value == null)
            {
                dgv.CurrentRow.Cells[kolumna].Value = "0,00";

                return;
            }


            if (czyNumeric(dgv.CurrentRow.Cells[kolumna].Value.ToString(), "d"))
            {
                double d = double.Parse(dgv.CurrentRow.Cells[kolumna].Value.ToString());

                if (d < 0)
                {
                    dgv.CurrentRow.Cells[kolumna].Value = ((-1) * d).ToString("N2");

                }
                else
                {
                    dgv.CurrentRow.Cells[kolumna].Value = d.ToString("N2");

                }
                dgv.CurrentRow.Cells[kolumna].Style.BackColor = Color.AliceBlue;
            }

            else
            {
                dgv.CurrentRow.Cells[kolumna].Style.BackColor = Color.Red;
                dgv.CurrentRow.Cells[kolumna].Value = "0,00";
            }
        }

        // formatowanie Int
        void formatowanieInt(DataGridView dgv, int kolumna)
        {

            if (dgv.Rows.Count == 1)
            {
                return;
            }

            if (dgv.CurrentRow.Cells[kolumna].Value == null)
            {
                dgv.CurrentRow.Cells[kolumna].Value =0;
                return;
            }


            if (czyNumeric(dgv.CurrentRow.Cells[kolumna].Value.ToString(), "n"))
            {
                int n = Int32.Parse(dgv.CurrentRow.Cells[kolumna].Value.ToString());
                if(n<0)
                {
                    dgv.CurrentRow.Cells[kolumna].Value = (-1) * n;
                }

               
                dgv.CurrentRow.Cells[kolumna].Style.BackColor = Color.AliceBlue;
            }

            else
            {
                dgv.CurrentRow.Cells[kolumna].Style.BackColor = Color.Red;
                dgv.CurrentRow.Cells[kolumna].Value = 0;

            }
        }

        //funckja sprawdzająca czy pole ma wartosc numeryczną lub NULL
        public bool czyNumeric(string s, string wartosc)
        {
            if (wartosc == "n")
            {

                int y;
                bool m = int.TryParse(s, out y);

                if (m == false)
                {
                    MessageBox.Show("nieprawidłowy format liczbowy", "bład", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return false;
                }

                else
                {

                    return true;
                }

            }

            else
            {
                double y;
                bool m = double.TryParse(s, out y);

                if (m == false)
                {
                    MessageBox.Show("nieprawidłowy format liczbowy", "bład", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return false;
                }

                else
                {

                    return true;
                }

            }
        }
        // - - -----------------------------------------------


        // sumowanie dowolnej kolumny z dgv
        double sumuj(DataGridView dgv, int kolumna)
        {
            double x = 0;

            for (int k = 0; k < dgv.Rows.Count - 1; k++)
            {
                if (dgv.Rows[k].Cells[kolumna].Value == null || czyNumeric(dgv.Rows[k].Cells[kolumna].Value.ToString(), "d") == false)
                {
                    return x=0;

                }

                x = double.Parse(dgv.Rows[k].Cells[kolumna].Value.ToString()) + x;


            }


            return x;


        }

        //  - - - KONIEC   funkcje techniczne ----------------------------------------------------------------
        

        //  TESTOWE -- -####################################################################################



        //duplikat funkcji wczytaj - pierwsze wczytanie kursu


        // drugie i kolejne wczytania kursu
        public void read_currency(string actual_file)
        {

            try
            {

           

                DataSet ds = new DataSet();
                ds.ReadXml("http://www.nbp.pl/kursy/xml/" + actual_file + ".xml");

                kursy.DataSource = ds.Tables[1];
          

                label13.Text = "Data publikacji: " + ds.Tables[0].Rows[0]["data_publikacji"].ToString();
                internetConncetion = true;
          

            }

            catch
            {
                MessageBox.Show ("Brak połączenia z internetem, brak tabeli NBP lub strona NBP nie odpowiada. Rozliczenie delegacji zagranicznej nie będzie mozliwe!");

                internetConncetion = false;
                tabeleZagraniczna.Visible = false;

            }
        }


       private void drukujPDF_Click(object sender, EventArgs e)
             {
                    if(tabeleKrajowa.Visible==true)
                    {

                        drukuj_krajowa();
                        MessageBox.Show("wydrukowano na pulpicie!");
                    }

                    else
                    {

                        drukuj_zagraniczna();
                        MessageBox.Show("wydrukowano na pulpicie!");
                    }

             }


        void drukuj_krajowa()
       {

           double sumaDelegacji = 0;


           int licznik = 0;
           string tekst = "";


           //EKSPORT JEZELI POLA IMIE NAZIWSKO ORAZ OKRES I W TABELACH ZOSTALY UZUPELNIONE WARTOSCI SA NIEPUSTE

           // DRUKOWANIE FORMULARZA
           //SCIEZKA ZAPISU

           string pdfpath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);



           DateTime dataWydruku = DateTime.Now;


           string dataWydruku1 = "   " + DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString() + DateTime.Now.Day.ToString() + "  " + DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString();

           //nazwa dokumentu:
           string nazwa_pliku;
           string rodzdel;



               nazwa_pliku = "ROZLICZENIE DEL.KRAJOWEJ_ZALICZKI" + "_" + dataWydruku1.ToString() + ".pdf";
               rodzdel = "ROZLICZENIE DELEGACJI KRAJOWEJ/ZALICZKI";
             




            Document doc = new Document();
            
           try
           {


               PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(pdfpath + "\\" + nazwa_pliku, FileMode.Create));

               doc.Open();



               // ################################################## //
               //DODANIE LOGO

               iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(imageList1.Images[0], System.Drawing.Imaging.ImageFormat.Png);
               logo.SetAbsolutePosition(410, 770);

               doc.Add(logo);

               

                // ################################################## //
                // CZCIONKI




                BaseFont bfTimes = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, false);
                BaseFont arial = BaseFont.CreateFont(@"C:\\WINDOWS\\Fonts\\arial.ttf", "iso-8859-2", BaseFont.EMBEDDED);

                iTextSharp.text.Font fontNaglowek = FontFactory.GetFont("arial", 14, iTextSharp.text.Font.BOLD, BaseColor.BLACK);

                iTextSharp.text.Font font14 = new iTextSharp.text.Font(arial, 14);


                iTextSharp.text.Font font13 = new iTextSharp.text.Font(arial, 13);
               iTextSharp.text.Font font11 = new iTextSharp.text.Font(arial, 11);
               iTextSharp.text.Font font10 = new iTextSharp.text.Font(arial, 10);
               iTextSharp.text.Font font21 = new iTextSharp.text.Font(arial, 10, iTextSharp.text.Font.BOLD);
               iTextSharp.text.Font font9 = new iTextSharp.text.Font(arial, 9);
               iTextSharp.text.Font font7 = new iTextSharp.text.Font(arial, 7);
               iTextSharp.text.Font font8 = new iTextSharp.text.Font(arial, 8);

               iTextSharp.text.Font font6 = new iTextSharp.text.Font(arial, 6);

               // ################################################## //
               // NAGŁÓWEK



               Phrase naglowek = new Phrase(rodzdel, font13);
               Phrase ver = new Phrase(wersja, font7);
                doc.Add(naglowek);

                doc.Add(Chunk.NEWLINE);
                doc.Add(ver);


                //spacja

                doc.Add(Chunk.NEWLINE);
               doc.Add(Chunk.NEWLINE);


               // ################################################## //
               //DODANIE PIERWSZEJ TABELI - nagłówek


               PdfPTable pdfTable_naglowek = new PdfPTable(2);
               //     PdfPCell cell = new PdfPCell(new Paragraph("RODZAJ DELEGACJI:  \tDELEGACJA" + " "+rodzajDel, font11));
               PdfPCell cell = new PdfPCell();

               pdfTable_naglowek.WidthPercentage = 55;
               float[] widths = new float[] { 40f, 90f };
               pdfTable_naglowek.SetWidths(widths);



               cell.Colspan = 2;
               pdfTable_naglowek.AddCell(cell);

               pdfTable_naglowek.AddCell(new Phrase("IMIĘ I NAZWISKO", font9));
               pdfTable_naglowek.AddCell(new Phrase(textBox1.Text.ToString(), font9));
               pdfTable_naglowek.AddCell(new Phrase("MPK", font9));
               pdfTable_naglowek.AddCell(new Phrase(textBox3.Text.ToString(), font9));
               pdfTable_naglowek.AddCell(new Phrase("ID PRACOWNIKA", font9));
               pdfTable_naglowek.AddCell(new Phrase(textBox2.Text.ToString(), font9));
               if (textBox6.Text!="")
                {
                    pdfTable_naglowek.AddCell(new Phrase("ID karty VISA", font9));
                    pdfTable_naglowek.AddCell(new Phrase(textBox6.Text.ToString(), font9));
                }
                        
                          
               pdfTable_naglowek.AddCell(new Phrase("OKRES ROZLICZENIOWY", font9));
               pdfTable_naglowek.AddCell(new Phrase(textBox5.Text, font9));
              // pdfTable_naglowek.AddCell(new Phrase("nr rach. bankowego", font9));
              // pdfTable_naglowek.AddCell(new Phrase(textBox4.Text, font9));


               string datawy;
               datawy = dataWydruku.ToString();
               pdfTable_naglowek.AddCell(new Phrase(datawy, font9));

               doc.Add(pdfTable_naglowek);


               doc.Add(Chunk.NEWLINE);


               // ################################################## //
               // TABLICA delegacje //

               PdfPTable pdfTableDelegacje = new PdfPTable(13);
               PdfPCell cell4 = new PdfPCell(new Paragraph("DELEGACJE", font21));
               pdfTableDelegacje.WidthPercentage = 85;

               pdfTableDelegacje.HorizontalAlignment = 0;


               float[] widths4 = new float[] { 0f, 45f, 30f, 45f, 30f, 65f, 85f, 10f, 10f, 10f, 35f, 35f, 35f };
               pdfTableDelegacje.SetWidths(widths4);




               cell4.Colspan = 13;
               pdfTableDelegacje.AddCell(cell4);

               pdfTableDelegacje.AddCell(new Phrase("", font9));
               pdfTableDelegacje.AddCell(new Phrase("DATA ROZP", font9));
               pdfTableDelegacje.AddCell(new Phrase("GODZ ROZP", font9));
               pdfTableDelegacje.AddCell(new Phrase("DATA ZAK", font9));
               pdfTableDelegacje.AddCell(new Phrase("GODZ ZAK", font9));
               pdfTableDelegacje.AddCell(new Phrase("RAZEM", font9));
               pdfTableDelegacje.AddCell(new Phrase("CEL I MIEJSCOWOŚĆ", font9));
               pdfTableDelegacje.AddCell(new Phrase("Ś", font9));
               pdfTableDelegacje.AddCell(new Phrase("O", font9));
               pdfTableDelegacje.AddCell(new Phrase("K", font9));
               pdfTableDelegacje.AddCell(new Phrase("DIETA", font9));
               pdfTableDelegacje.AddCell(new Phrase("KONTO", font9));
               pdfTableDelegacje.AddCell(new Phrase("IN. - AUF", font9));

               licznik = delegacjeKrajowa.Rows.Count;



               if (licznik != 1)
               {
                   foreach (DataGridViewRow row in delegacjeKrajowa.Rows)
                   {
                       foreach (DataGridViewCell kom1 in row.Cells)
                       {


                           if (kom1.Value != null)
                           {

                               tekst = kom1.Value.ToString();

                               // formatowanie INNENAUFTRAGÓW //
                               if (kom1.ColumnIndex == 12)
                               {
                                   tekst = tekst.Substring(0, 5);

                               }


                               pdfTableDelegacje.AddCell(new Phrase(tekst, font8));
                           }

                           else
                           {

                               tekst = "-";

                               pdfTableDelegacje.AddCell((new Phrase(tekst, font8)));
                           }


                       }
                   }

                   PdfPCell suma = new PdfPCell(new Paragraph("RAZEM:  " + label35.Text + "  PLN           ", font21));
                   suma.Colspan = 13;
                   suma.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                   pdfTableDelegacje.AddCell(suma);

                   sumaDelegacji = double.Parse(label35.Text.ToString()) + sumaDelegacji;

                   doc.Add(pdfTableDelegacje);
               }

               else
               {


                   PdfPCell cellblad4 = new PdfPCell(new Paragraph(" -  nie zadeklarowano -  ", font11));
                   cellblad4.Colspan = 13;
                   cellblad4.HorizontalAlignment = 1;

                   pdfTableDelegacje.AddCell(cellblad4);

                   doc.Add(pdfTableDelegacje);


               }


               doc.Add(Chunk.NEWLINE);
               doc.Add(Chunk.NEWLINE);

               // ################################################## //
               //DODANIE DRUGIEJ TABELI - noclegi


               PdfPTable pdfTableNoclegi = new PdfPTable(6);
               PdfPCell cell2 = new PdfPCell(new Paragraph("NOCLEGI", font21));
               pdfTableNoclegi.WidthPercentage = 70;

               pdfTableNoclegi.HorizontalAlignment = 0;


               float[] widths2 = new float[] { 0f, 70f, 20f, 15f, 15f, 15f };
               pdfTableNoclegi.SetWidths(widths2);




               cell2.Colspan = 6;
               pdfTableNoclegi.AddCell(cell2);

               pdfTableNoclegi.AddCell(new Phrase("", font9));
               pdfTableNoclegi.AddCell(new Phrase("NOCLEG", font9));
               pdfTableNoclegi.AddCell(new Phrase("LICZBA", font9));
               pdfTableNoclegi.AddCell(new Phrase("KWOTA", font9));
               pdfTableNoclegi.AddCell(new Phrase("KONTO", font9));
               pdfTableNoclegi.AddCell(new Phrase("IN.-AUF", font9));

               licznik = noclegiKrajowa.Rows.Count;



               if (licznik != 1)
               {
                   foreach (DataGridViewRow row in noclegiKrajowa.Rows)
                   {
                       foreach (DataGridViewCell kom1 in row.Cells)
                       {


                           if (kom1.Value != null)
                           {

                               tekst = kom1.Value.ToString();

                             // formatowanie INNENAUFTRAGÓW //
                               if(kom1.ColumnIndex==5)
                              {
                                 tekst = tekst.Substring(0,5);

                               }

                               pdfTableNoclegi.AddCell((new Phrase(tekst, font8)));
                           }


                           else
                           {

                               tekst = "-";

                               pdfTableNoclegi.AddCell((new Phrase(tekst, font8)));
                           }

                       }
                   }




                   PdfPCell suma = new PdfPCell(new Paragraph("IL. NOCL.: " + label17.Text + ", " + " SUMA: " + label18.Text + " PLN", font21));
                   suma.Colspan = 6;
                   suma.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                   pdfTableNoclegi.AddCell(suma);
                   sumaDelegacji = double.Parse(label18.Text.ToString()) + sumaDelegacji;
                   doc.Add(pdfTableNoclegi);
               }

               else
               {
                   PdfPCell cellblad1 = new PdfPCell(new Paragraph(" -  nie zadeklarowano - ", font11));
                   cellblad1.Colspan = 6;
                   cellblad1.HorizontalAlignment = 1;

                   pdfTableNoclegi.AddCell(cellblad1);

                   doc.Add(pdfTableNoclegi);


               }


               doc.Add(Chunk.NEWLINE);
               // ################################################## //
               //DODANIE trzeciej TABELI - koszty dodatkowe

               PdfPTable pdfTableDodatkowe = new PdfPTable(5);
               PdfPCell cell3 = new PdfPCell(new Paragraph("KOSZTY DODATKOWE", font21));
               pdfTableDodatkowe.WidthPercentage = 80;

               pdfTableDodatkowe.HorizontalAlignment = 0;


               float[] widths3 = new float[] { 0f, 68f, 15f, 15f, 15f };
               pdfTableDodatkowe.SetWidths(widths3);




               cell3.Colspan = 5;
               pdfTableDodatkowe.AddCell(cell3);

               pdfTableDodatkowe.AddCell(new Phrase("", font9));
               pdfTableDodatkowe.AddCell(new Phrase("RODZAJ", font9));
               pdfTableDodatkowe.AddCell(new Phrase("LICZBA", font9));
               pdfTableDodatkowe.AddCell(new Phrase("KONTO", font9));
               pdfTableDodatkowe.AddCell(new Phrase("INNEN-AUF.", font9));

               licznik = dodatkoweKrajowa.Rows.Count;



               if (licznik != 1)
               {
                   foreach (DataGridViewRow row in dodatkoweKrajowa.Rows)
                   {
                       foreach (DataGridViewCell kom1 in row.Cells)
                       {


                           if (kom1.Value != null)
                           {

                               tekst = kom1.Value.ToString();

                               // formatowanie INNENAUFTRAGÓW //
                               if (kom1.ColumnIndex == 4)
                               {
                                   tekst = tekst.Substring(0, 5);

                               }

                               pdfTableDodatkowe.AddCell((new Phrase(tekst, font8)));
                           }

                           else
                           {

                               tekst = "-";

                               pdfTableDodatkowe.AddCell((new Phrase(tekst, font8)));
                           }


                       }
                   }

                   PdfPCell suma = new PdfPCell(new Paragraph("RAZEM:  " + label20.Text + "  PLN           ", font21));
                   suma.Colspan = 5;
                   suma.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                   pdfTableDodatkowe.AddCell(suma);
                   sumaDelegacji = double.Parse(label20.Text.ToString()) + sumaDelegacji;
                   doc.Add(pdfTableDodatkowe);
               }

               else
               {
                   PdfPCell cellblad2 = new PdfPCell(new Paragraph(" -  nie zadeklarowano - ", font11));
                   cellblad2.Colspan = 5;
                   cellblad2.HorizontalAlignment = 1;

                   pdfTableDodatkowe.AddCell(cellblad2);

                   doc.Add(pdfTableDodatkowe);


               }


               doc.Add(Chunk.NEWLINE);




               // ################################################## //
               //DODANIE CZWARTEJ TABELI - koszty samochodu

               PdfPTable pdfTableKosztySamochodu = new PdfPTable(5);
               PdfPCell cell1 = new PdfPCell(new Paragraph("KOSZTY SAMOCHODU", font21));
               pdfTableKosztySamochodu.WidthPercentage = 70;

               pdfTableKosztySamochodu.HorizontalAlignment = 0;


               float[] widths1 = new float[] { 0f, 100f, 15f, 15f, 25f };
               pdfTableKosztySamochodu.SetWidths(widths1);




               cell1.Colspan = 5;
               pdfTableKosztySamochodu.AddCell(cell1);

               pdfTableKosztySamochodu.AddCell(new Phrase("", font9));
               pdfTableKosztySamochodu.AddCell(new Phrase("RODZAJ", font9));
               pdfTableKosztySamochodu.AddCell(new Phrase("KWOTA", font9));
               pdfTableKosztySamochodu.AddCell(new Phrase("KONTO", font9));
               pdfTableKosztySamochodu.AddCell(new Phrase("INN-AUF", font9));
               licznik = samochoduKrajowa.Rows.Count;



               if (licznik != 1)
               {
                   foreach (DataGridViewRow row in samochoduKrajowa.Rows)
                   {
                       foreach (DataGridViewCell kom1 in row.Cells)
                       {


                           if (kom1.Value != null)
                           {

                               tekst = kom1.Value.ToString();

                               // formatowanie INNENAUFTRAGÓW //
                               if (kom1.ColumnIndex == 4)
                               {
                                   tekst = tekst.Substring(0, 5);

                               }

                               pdfTableKosztySamochodu.AddCell((new Phrase(tekst, font8)));
                           }

                           else
                           {

                               tekst = "-";

                               pdfTableKosztySamochodu.AddCell((new Phrase(tekst, font8)));
                           }

                       }
                   }

                   PdfPCell suma = new PdfPCell(new Paragraph("RAZEM:  " + label22.Text + "  PLN           ", font21));
                   suma.Colspan = 5;
                   suma.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                   pdfTableKosztySamochodu.AddCell(suma);
                   sumaDelegacji = double.Parse(label22.Text.ToString()) + sumaDelegacji;
                   doc.Add(pdfTableKosztySamochodu);
               }

               else
               {
                   PdfPCell cellblad = new PdfPCell(new Paragraph(" -  nie zadeklarowano - ", font11));
                   cellblad.Colspan = 5;
                   cellblad.HorizontalAlignment = 1;
                   pdfTableKosztySamochodu.AddCell(cellblad);

                   doc.Add(pdfTableKosztySamochodu);


               }

               doc.Add(Chunk.NEWLINE);



               //DODAWANIE stopki

               doc.Add(Chunk.NEWLINE);

               Phrase podusmowanie = new Phrase("RAZEM: " + sumaDelegacji.ToString("N2") + " PLN", font21);
               doc.Add(podusmowanie);
               doc.Add(Chunk.NEWLINE);
               doc.Add(Chunk.NEWLINE);
               doc.Add(Chunk.NEWLINE);

               Phrase stopka1 = new Phrase("podpis pracownika" + "                     " + "podpis przełożonego" + "                  " + "podpis osoby ostatecznie akceptującej", font6);
               doc.Add(stopka1);

               //spacja

               doc.Add(Chunk.NEWLINE);
               doc.Add(Chunk.NEWLINE);

               Phrase stopka2 = new Phrase(".............................." + "                   " + " .................................." + "               " + " .................................................", font6);
               doc.Add(stopka2);

               //spacja

               doc.Add(Chunk.NEWLINE);
               doc.Add(Chunk.NEWLINE);
               doc.Add(Chunk.NEWLINE);

               Phrase stopka3 = new Phrase("DATA DOKUMENTU:" + dataWydruku.ToString(), font6);
               doc.Add(stopka3);
               //spacja



               doc.Add(Chunk.NEWLINE);
           }




           catch ( Exception ex)
           {
               MessageBox.Show("wystapił błąd " + ex.ToString());
           }

           finally
           {



               doc.Close();
           }
       }

        void drukuj_zagraniczna()
        {

            double sumaDelegacji = 0;


            int licznik = 0;
            string tekst = "";


            //EKSPORT JEZELI POLA IMIE NAZIWSKO ORAZ OKRES I W TABELACH ZOSTALY UZUPELNIONE WARTOSCI SA NIEPUSTE

            // DRUKOWANIE FORMULARZA
            //SCIEZKA ZAPISU

            string pdfpath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);



            DateTime dataWydruku = DateTime.Now;


            string dataWydruku1 = "   " + DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString() + DateTime.Now.Day.ToString() + "  " + DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString();

            //nazwa dokumentu:
            string nazwa_pliku;

            string rodzdel;



            nazwa_pliku = "ROZLICZENIE DEL.ZAGR_ZALICZKI" + dataWydruku1.ToString() + ".pdf";
            rodzdel = "ROZLICZENIE DELEGACJI ZAGRANICZNEJ/ZALICZKI";
       





            Document doc = new Document();

            try
            {


                PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(pdfpath + "\\" + nazwa_pliku, FileMode.Create));

                doc.Open();



                // ################################################## //
                //DODANIE LOGO

                iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(imageList1.Images[0], System.Drawing.Imaging.ImageFormat.Png);
                logo.SetAbsolutePosition(410, 770);

           
                


                doc.Add(logo);

                // ################################################## //
                // CZCIONKI

                BaseFont bfTimes = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, false);
                BaseFont arial = BaseFont.CreateFont(@"C:\\WINDOWS\\Fonts\\arial.ttf", "iso-8859-2", BaseFont.EMBEDDED);

                iTextSharp.text.Font fontNaglowek = FontFactory.GetFont("arial", 14, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
                
                iTextSharp.text.Font font14 = new iTextSharp.text.Font(arial, 14);

                iTextSharp.text.Font font13 = new iTextSharp.text.Font(arial, 13);
                iTextSharp.text.Font font11 = new iTextSharp.text.Font(arial, 11);
                iTextSharp.text.Font font10 = new iTextSharp.text.Font(arial, 10);
                iTextSharp.text.Font font9 = new iTextSharp.text.Font(arial, 9);
                iTextSharp.text.Font font7 = new iTextSharp.text.Font(arial, 7);
                iTextSharp.text.Font font8 = new iTextSharp.text.Font(arial, 8);
                iTextSharp.text.Font font21 = new iTextSharp.text.Font(arial, 10, iTextSharp.text.Font.BOLD);
             

                iTextSharp.text.Font font6 = new iTextSharp.text.Font(arial, 6);

                // ################################################## //
                // NAGŁÓWEK



                Phrase naglowek = new Phrase(rodzdel,font13);
                Phrase ver = new Phrase(wersja, font7);
                doc.Add(naglowek);

                doc.Add(Chunk.NEWLINE);
                doc.Add(ver);
                //spacja

                doc.Add(Chunk.NEWLINE);
                doc.Add(Chunk.NEWLINE);


                // ################################################## //
                //DODANIE PIERWSZEJ TABELI - nagłówek


                PdfPTable pdfTable_naglowek = new PdfPTable(2);
                //     PdfPCell cell = new PdfPCell(new Paragraph("RODZAJ DELEGACJI:  \tDELEGACJA" + " "+rodzajDel, font11));
                PdfPCell cell = new PdfPCell();

                pdfTable_naglowek.WidthPercentage = 55;
                float[] widths = new float[] { 40f, 90f };
                pdfTable_naglowek.SetWidths(widths);



                cell.Colspan = 2;
                pdfTable_naglowek.AddCell(cell);

                pdfTable_naglowek.AddCell(new Phrase("IMIĘ I NAZWISKO", font9));
                pdfTable_naglowek.AddCell(new Phrase(textBox1.Text.ToString(), font9));
                pdfTable_naglowek.AddCell(new Phrase("MPK", font9));
                pdfTable_naglowek.AddCell(new Phrase(textBox3.Text.ToString(), font9));
                pdfTable_naglowek.AddCell(new Phrase("ID PRACOWNIKA", font9));
                pdfTable_naglowek.AddCell(new Phrase(textBox2.Text.ToString(), font9));
                if (textBox6.Text != "")
                {
                    pdfTable_naglowek.AddCell(new Phrase("ID karty VISA", font9));
                    pdfTable_naglowek.AddCell(new Phrase(textBox6.Text.ToString(), font9));
                }
                pdfTable_naglowek.AddCell(new Phrase("OKRES ROZLICZENIOWY", font9));
                pdfTable_naglowek.AddCell(new Phrase(textBox5.Text, font9));
               // pdfTable_naglowek.AddCell(new Phrase("nr rach. bankowego", font9));
              //  pdfTable_naglowek.AddCell(new Phrase(textBox4.Text, font9));


                string datawy;
                datawy = dataWydruku.ToString();
                pdfTable_naglowek.AddCell(new Phrase(datawy, font9));

                doc.Add(pdfTable_naglowek);



                //PODKRESLENIE
                //Paragraph p = new Paragraph(new Chunk(new iTextSharp.text.pdf.draw.LineSeparator(2.0F, 100.0F, BaseColor.ORANGE, Element.ALIGN_LEFT, 1)));
                //doc.Add(p);

                doc.Add(Chunk.NEWLINE);

                //okreslenie jak został wczytany kurs

                Phrase obramowanie = new Phrase("-------------------------------------------------------------------------------------------------------", font6);
                doc.Add(obramowanie);

                doc.Add(Chunk.NEWLINE);

                if(Form2.kursNapodstawiezaliczki==true)
                {

                    Phrase opis0 = new Phrase("ROZLICZENIE Z DNIA: " + Form2.data1.ToString(), font6);
                    doc.Add(opis0);
                    doc.Add(Chunk.NEWLINE);
                    doc.Add(Chunk.NEWLINE);
                    Phrase opis = new Phrase("KURS NA PODSTAWIE WPŁ. ZALICZKI Z DNIA: " + Form2.data+ " W WALUCIE:  "+ Form2.zalKsiegowosc,font6);
                    doc.Add(opis);
                    doc.Add(Chunk.NEWLINE);

                }

                else
                {
                    Phrase opis = new Phrase("ROZLICZENIE Z DNIA: " + Form2.data1.ToString() + " KURS  Z DNIA: " + Form2.data.ToString(),font6 );
                    doc.Add(opis);
                    doc.Add(Chunk.NEWLINE);

                }

                Phrase obramowanie1= new Phrase("--------------------------------------------------------------------------------------------------------", font6);
                doc.Add(obramowanie1);

                // ################################################## //
                // TABLICA delegacje //

                PdfPTable pdfTableDelegacje = new PdfPTable(17);
                PdfPCell cell4 = new PdfPCell(new Paragraph("DELEGACJE", font21));
                pdfTableDelegacje.WidthPercentage = 95;

                pdfTableDelegacje.HorizontalAlignment = 0;


                float[] widths4 = new float[] { 0f, 70f, 40f, 70f, 40f, 85f, 100f, 65f, 35f, 50f, 15f, 15f, 15f, 65f,65f,50f,45f };
                pdfTableDelegacje.SetWidths(widths4);




                cell4.Colspan = 17;
                pdfTableDelegacje.AddCell(cell4);

                pdfTableDelegacje.AddCell(new Phrase("", font9));
                pdfTableDelegacje.AddCell(new Phrase("DATA ROZP", font9));
                pdfTableDelegacje.AddCell(new Phrase("GODZ ROZP", font9));
                pdfTableDelegacje.AddCell(new Phrase("DATA ZAK", font9));
                pdfTableDelegacje.AddCell(new Phrase("GODZ ZAK", font9));
                pdfTableDelegacje.AddCell(new Phrase("RAZEM", font9));
                pdfTableDelegacje.AddCell(new Phrase("CEL I MIEJSCOWOŚĆ", font9));
                pdfTableDelegacje.AddCell(new Phrase("KRAJ", font9));
                pdfTableDelegacje.AddCell(new Phrase("WALUTA", font9));
                pdfTableDelegacje.AddCell(new Phrase("KURS", font9));
                pdfTableDelegacje.AddCell(new Phrase("Ś", font9));
                pdfTableDelegacje.AddCell(new Phrase("O", font9));
                pdfTableDelegacje.AddCell(new Phrase("K", font9));
                pdfTableDelegacje.AddCell(new Phrase("DIETA W WAL.", font9));
                pdfTableDelegacje.AddCell(new Phrase("DIETA", font9));
                pdfTableDelegacje.AddCell(new Phrase("KONTO", font9));
                pdfTableDelegacje.AddCell(new Phrase("IN. - AUF", font9));

                licznik = delegacjeZagraniczna.Rows.Count;



                if (licznik != 1)
                {
                    foreach (DataGridViewRow row in delegacjeZagraniczna.Rows)
                    {
                        foreach (DataGridViewCell kom1 in row.Cells)
                        {


                            if (kom1.Value != null)
                            {

                                tekst = kom1.Value.ToString();




                                pdfTableDelegacje.AddCell(new Phrase(tekst, font8));
                            }

                            else
                            {

                                tekst = "-";

                                pdfTableDelegacje.AddCell((new Phrase(tekst, font8)));
                            }


                        }
                    }

                    PdfPCell suma = new PdfPCell(new Paragraph("RAZEM:  " + label52.Text + "  PLN           ", font21));
                    suma.Colspan = 17;
                    suma.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                    pdfTableDelegacje.AddCell(suma);

                    sumaDelegacji = double.Parse(label52.Text.ToString()) + sumaDelegacji;

                    doc.Add(pdfTableDelegacje);
                }

                else
                {


                    PdfPCell cellblad4 = new PdfPCell(new Paragraph(" -  nie zadeklarowano -  ", font11));
                    cellblad4.Colspan = 17;
                    cellblad4.HorizontalAlignment = 1;

                    pdfTableDelegacje.AddCell(cellblad4);

                    doc.Add(pdfTableDelegacje);


                }


                doc.Add(Chunk.NEWLINE);

                // ################################################## //
                //DODANIE DRUGIEJ TABELI - noclegi


                PdfPTable pdfTableNoclegi = new PdfPTable(9);
                PdfPCell cell2 = new PdfPCell(new Paragraph("NOCLEGI", font21));
                pdfTableNoclegi.WidthPercentage = 70;

                pdfTableNoclegi.HorizontalAlignment = 0;


                float[] widths2 = new float[] { 0f, 70f, 20f, 15f, 15f, 15f, 15f, 15f, 15f };
                pdfTableNoclegi.SetWidths(widths2);




                cell2.Colspan = 9;
                pdfTableNoclegi.AddCell(cell2);

                pdfTableNoclegi.AddCell(new Phrase("", font9));
                pdfTableNoclegi.AddCell(new Phrase("NOCLEG", font9));
                pdfTableNoclegi.AddCell(new Phrase("LICZBA", font9));
                pdfTableNoclegi.AddCell(new Phrase("KWOTA", font9));
                pdfTableNoclegi.AddCell(new Phrase("WALUTA", font9));
                pdfTableNoclegi.AddCell(new Phrase("KURS", font9));
                pdfTableNoclegi.AddCell(new Phrase("KWOTA W PLN", font9));
                pdfTableNoclegi.AddCell(new Phrase("KONTO", font9));
                pdfTableNoclegi.AddCell(new Phrase("IN.-AUF", font9));

                licznik =  noclegiZagraniczna.Rows.Count;



                if (licznik != 1)
                {
                    foreach (DataGridViewRow row in noclegiZagraniczna.Rows)
                    {
                        foreach (DataGridViewCell kom1 in row.Cells)
                        {


                            if (kom1.Value != null)
                            {

                                tekst = kom1.Value.ToString();

                                pdfTableNoclegi.AddCell((new Phrase(tekst, font8)));
                            }


                            else
                            {

                                tekst = "-";

                                pdfTableNoclegi.AddCell((new Phrase(tekst, font8)));
                            }

                        }
                    }




                    PdfPCell suma = new PdfPCell(new Paragraph("IL. NOCL.:  " + label25.Text + ", " + "SUMA: " + label27.Text + " PLN", font21));
                    suma.Colspan = 9;
                    suma.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                    pdfTableNoclegi.AddCell(suma);
                    sumaDelegacji = double.Parse(label27.Text.ToString()) + sumaDelegacji;
                    doc.Add(pdfTableNoclegi);
                }

                else
                {
                    PdfPCell cellblad1 = new PdfPCell(new Paragraph(" -  nie zadeklarowano - ", font11));
                    cellblad1.Colspan = 9;
                    cellblad1.HorizontalAlignment = 1;

                    pdfTableNoclegi.AddCell(cellblad1);

                    doc.Add(pdfTableNoclegi);


                }


                doc.Add(Chunk.NEWLINE);
                // ################################################## //
                //DODANIE trzeciej TABELI - koszty dodatkowe

                PdfPTable pdfTableDodatkowe = new PdfPTable(8);
                PdfPCell cell3 = new PdfPCell(new Paragraph("KOSZTY DODATKOWE", font21));
                pdfTableDodatkowe.WidthPercentage = 85;

                pdfTableDodatkowe.HorizontalAlignment = 0;


                float[] widths3 = new float[] { 0f, 68f, 15f, 15f, 15f, 15f, 15f, 15f };
                pdfTableDodatkowe.SetWidths(widths3);




                cell3.Colspan = 8;
                pdfTableDodatkowe.AddCell(cell3);

                pdfTableDodatkowe.AddCell(new Phrase("", font9));
                pdfTableDodatkowe.AddCell(new Phrase("RODZAJ", font9));
                pdfTableDodatkowe.AddCell(new Phrase("KWOTA", font9));
                pdfTableDodatkowe.AddCell(new Phrase("WALUTA", font9));
                pdfTableDodatkowe.AddCell(new Phrase("KURS", font9));
                pdfTableDodatkowe.AddCell(new Phrase("KWOTA W PLN", font9));
                pdfTableDodatkowe.AddCell(new Phrase("KONTO", font9));
                pdfTableDodatkowe.AddCell(new Phrase("INNEN-AUF.", font9));

                licznik = dodatkoweZagraniczna.Rows.Count;



                if (licznik != 1)
                {
                    foreach (DataGridViewRow row in dodatkoweZagraniczna.Rows)
                    {
                        foreach (DataGridViewCell kom1 in row.Cells)
                        {


                            if (kom1.Value != null)
                            {

                                tekst = kom1.Value.ToString();

                                pdfTableDodatkowe.AddCell((new Phrase(tekst, font8)));
                            }

                            else
                            {

                                tekst = "-";

                                pdfTableDodatkowe.AddCell((new Phrase(tekst, font8)));
                            }


                        }
                    }

                    PdfPCell suma = new PdfPCell(new Paragraph("RAZEM:  " + label44.Text + "  PLN           ", font21));
                    suma.Colspan = 8;
                    suma.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                    pdfTableDodatkowe.AddCell(suma);
                    sumaDelegacji = double.Parse(label44.Text.ToString()) + sumaDelegacji;
                    doc.Add(pdfTableDodatkowe);
                }

                else
                {
                    PdfPCell cellblad2 = new PdfPCell(new Paragraph(" -  nie zadeklarowano - ", font11));
                    cellblad2.Colspan = 8;
                    cellblad2.HorizontalAlignment = 1;

                    pdfTableDodatkowe.AddCell(cellblad2);

                    doc.Add(pdfTableDodatkowe);


                }


                doc.Add(Chunk.NEWLINE);




                // ################################################## //
                //DODANIE CZWARTEJ TABELI - koszty samochodu

                PdfPTable pdfTableKosztySamochodu = new PdfPTable(8);
                PdfPCell cell1 = new PdfPCell(new Paragraph("KOSZTY SAMOCHODU", font21));
                pdfTableKosztySamochodu.WidthPercentage = 75;

                pdfTableKosztySamochodu.HorizontalAlignment = 0;


                float[] widths1 = new float[] { 0f, 100f, 25f, 15f, 25f, 25f, 20f, 25f };
                pdfTableKosztySamochodu.SetWidths(widths1);




                cell1.Colspan = 8;
                pdfTableKosztySamochodu.AddCell(cell1);

                pdfTableKosztySamochodu.AddCell(new Phrase("", font9));
                pdfTableKosztySamochodu.AddCell(new Phrase("RODZAJ", font9));
                pdfTableKosztySamochodu.AddCell(new Phrase("KWOTA", font9));
                pdfTableKosztySamochodu.AddCell(new Phrase("WALUTA", font9));
                pdfTableKosztySamochodu.AddCell(new Phrase("KURS", font9));
                pdfTableKosztySamochodu.AddCell(new Phrase("KWOTA W PLN", font9));
                pdfTableKosztySamochodu.AddCell(new Phrase("KONTO", font9));
                pdfTableKosztySamochodu.AddCell(new Phrase("INN-AUF", font9));
                licznik = samochoduZagraniczna.Rows.Count;



                if (licznik != 1)
                {
                    foreach (DataGridViewRow row in samochoduZagraniczna.Rows)
                    {
                        foreach (DataGridViewCell kom1 in row.Cells)
                        {


                            if (kom1.Value != null)
                            {

                                tekst = kom1.Value.ToString();

                                pdfTableKosztySamochodu.AddCell((new Phrase(tekst, font8)));
                            }

                            else
                            {

                                tekst = "-";

                                pdfTableKosztySamochodu.AddCell((new Phrase(tekst, font8)));
                            }

                        }
                    }

                    PdfPCell suma = new PdfPCell(new Paragraph("RAZEM:  " + label1.Text + "  PLN           ", font21));
                    suma.Colspan = 8;
                    suma.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                    pdfTableKosztySamochodu.AddCell(suma);
                    sumaDelegacji = double.Parse(label1.Text.ToString()) + sumaDelegacji;
                    doc.Add(pdfTableKosztySamochodu);
                }

                else
                {
                    PdfPCell cellblad = new PdfPCell(new Paragraph(" -  nie zadeklarowano - ", font11));
                    cellblad.Colspan = 8;
                    cellblad.HorizontalAlignment = 1;
                    pdfTableKosztySamochodu.AddCell(cellblad);

                    doc.Add(pdfTableKosztySamochodu);


                }

                doc.Add(Chunk.NEWLINE);



                //DODAWANIE stopki

                doc.Add(Chunk.NEWLINE);

                Phrase podusmowanie = new Phrase("RAZEM: " + sumaDelegacji.ToString("N2") + " PLN", font21);
                doc.Add(podusmowanie);
                doc.Add(Chunk.NEWLINE);
                doc.Add(Chunk.NEWLINE);

                Phrase deklaracja= new Phrase("Wyrażam zgodę na rozliczenie delegacji zagranicznej w walucie polskiej.", font6);
                doc.Add(deklaracja);

                doc.Add(Chunk.NEWLINE);


                Phrase stopka1 = new Phrase("podpis pracownika" + "                     " + "podpis przełożonego" + "                  " + "podpis osoby ostatecznie akceptującej", font6);
                doc.Add(stopka1);

                //spacja

                doc.Add(Chunk.NEWLINE);
                doc.Add(Chunk.NEWLINE);

                Phrase stopka2 = new Phrase(".............................." + "                   " + " .................................." + "               " + " .................................................", font6);
                doc.Add(stopka2);

                //spacja

                doc.Add(Chunk.NEWLINE);
                doc.Add(Chunk.NEWLINE);
                doc.Add(Chunk.NEWLINE);

                Phrase stopka3 = new Phrase("DATA DOKUMENTU:" + dataWydruku.ToString(), font6);
                doc.Add(stopka3);
                //spacja

                doc.Add(Chunk.NEWLINE);
            }




            catch
            {
                MessageBox.Show("wystapił błąd");
            }

            finally
            {



                doc.Close();
            }
        }

       private void zapisdoXML_Click(object sender, EventArgs e)
             {

                 string sciezka = "\\";


                 if (noclegiKrajowa.Visible == true)
                 {

                  
                     sciezka = "\\saved\\noclegiKrajowa.xml";
                     zapisTabeliXML(noclegiKrajowa,sciezka);

                 }

                 if (dodatkoweKrajowa.Visible == true)
                 {
                     sciezka = "\\saved\\dodatkoweKrajowa.xml";
                     zapisTabeliXML(dodatkoweKrajowa, sciezka);


                 }

                 if (samochoduKrajowa.Visible == true)
                 {
                     sciezka = "\\saved\\samochoduKrajowa.xml";
                     zapisTabeliXML(samochoduKrajowa, sciezka);

                 }

                 if (delegacjeKrajowa.Visible == true)
                 {
                     sciezka = "\\saved\\delegacjeKrajowa.xml";
                     zapisTabeliXML(delegacjeKrajowa, sciezka);

                 }


                 if (noclegiZagraniczna.Visible == true)
                 {
                     sciezka = "\\saved\\noclegiZagraniczna.xml";
                     zapisTabeliXML(noclegiZagraniczna, sciezka);

                 }

                 if (dodatkoweZagraniczna.Visible == true)
                 {
                     sciezka = "\\saved\\dodatkoweZagraniczna.xml";
                     zapisTabeliXML(dodatkoweZagraniczna, sciezka);

                 }

                 if (samochoduZagraniczna.Visible == true)
                 {
                     sciezka = "\\saved\\samochoduZagraniczna.xml";
                     zapisTabeliXML(samochoduZagraniczna, sciezka);

                 }

                 if (delegacjeZagraniczna.Visible == true)
                 {
                     sciezka = "\\saved\\delegacjeZagraniczna.xml";
                     zapisTabeliXML(delegacjeZagraniczna, sciezka);

                 }

             

                
                }

      private void wczytaj_tabeleXML(DataGridView dgv)
       {


           //kasowanie aktualnej tabeli

           dgv.Rows.Clear();

          // sprawdzenie czy sa zapisane jakiekolwiek wartosci
          //jezeli plik jest pusty, to nie zostanie wczytany

           string plik = dgv.Name.ToString() + ".xml";


           if (new FileInfo("saved\\" + plik).Length == 0)
           {
               return;

           }

           // wczytanie tabeli

           else
           {

               DataSet ds = new DataSet();

               try
               {


                   ds.ReadXml("saved\\" + plik);



                   for (int k = 0; k < ds.Tables.Count; k++)
                   {
                       DataGridViewRow row = (DataGridViewRow)dgv.Rows[k].Clone();

                       dgv.Rows.Add(row);

                       for (int j = 1; j < columnCount(dgv); j++)
                       {
                           if (dgv.Rows[k].Cells[j].GetType() == typeof(DataGridViewComboBoxCell))
                           {


                               DataGridViewComboBoxCell comboCell = (DataGridViewComboBoxCell)dgv.Rows[k].Cells[j];
                               if (ds.Tables[k].Rows[0][j].ToString() == "puste")
                               {
                                   dgv.Rows[k].Cells[j].Value = null;

                               }

                               else
                               {
                                   int a = int.Parse(ds.Tables[k].Rows[0][j].ToString());

                                   dgv.Rows[k].Cells[j].Value = comboCell.Items[a];

                               }
                           }


                           else
                           {
                               dgv.Rows[k].Cells[j].Value = ds.Tables[k].Rows[0][j].ToString();

                           }



                       }



                   }




                   ds.Clear();



               }
               catch
               {

                   return;
               }

           }
           
       }


       private void readLastfile_Click(object sender, EventArgs e)
             {
                 wczytaj_tabeleXML(noclegiKrajowa);
                 wczytaj_tabeleXML(dodatkoweKrajowa);
                 wczytaj_tabeleXML(samochoduKrajowa);
                 wczytaj_tabeleXML(delegacjeKrajowa);
         
             }

        

        void zapisTabeliXML( DataGridView dgv, string sciezka)
       {
           
         
            if (dgv.Rows.Count==1)
            {

                System.IO.File.WriteAllText(@System.IO.Directory.GetCurrentDirectory().ToString() + sciezka, string.Empty);
                return;
            }

           XmlTextWriter writer = new XmlTextWriter(@System.IO.Directory.GetCurrentDirectory().ToString() + sciezka, System.Text.Encoding.UTF8);


           writer.WriteStartDocument(true);
           writer.Formatting = Formatting.Indented;
           writer.Indentation = 2;

           writer.WriteStartElement("tabele");

           int b = columnCount(dgv);


           for (int k = 0; k < dgv.Rows.Count - 1; k++)
           {
        
               writer.WriteStartElement("wiersz", k.ToString());


               for (int j = 0; j < b; j++)
               {


                   writer.WriteStartElement("cell" +j);

                   if (dgv.Rows[k].Cells[j].Value == null)
                   {
                       writer.WriteString("puste");


                   }


                   else
                   {

                       if (dgv.Rows[k].Cells[j].GetType() == typeof(DataGridViewComboBoxCell))
                       {
                           DataGridViewComboBoxCell comboCell = (DataGridViewComboBoxCell)dgv.Rows[k].Cells[j];
                           writer.WriteString(comboCell.Items.IndexOf(comboCell.Value).ToString());
                       }

                       else
                       {
                           writer.WriteString(dgv.Rows[k].Cells[j].Value.ToString());

                       }

                 
                   }

                   writer.WriteEndElement();
               }




               writer.WriteEndElement();

           }


           writer.WriteEndElement();

           writer.WriteEndDocument();
           writer.Close();

           progressBar2.Visible = true;
           progressBar2.Value = 50;

           for (int i = 0; i < 400; i++)
           {

               Thread.Sleep(1);
               progressBar2.Increment(3);
           }

           progressBar2.Visible = false;
        
       }


        int columnCount( DataGridView dgv)
        {
            
            int x = 0;

            foreach( DataGridViewColumn col in dgv.Columns)
            
            {
                if (col.Visible == true)
                {
                    x++;
                }

            }
            return x;

        }



        private void wprowadzDelegacje_Click(object sender, EventArgs e)
        {

            if (textBox17.Text.Length < 6)
            {

                MessageBox.Show("Wprowadź cel delegacji. min 7 znaków ", "CEL DELEGACJI", MessageBoxButtons.OK, MessageBoxIcon.Error);

                return;
            }



            //wartosc diety
            double dieta = 0;
          

            //wprowadzenie rekordu delegacji


            string dataRozp = dateTimePicker1.Value.ToString("dd-MM-yyyy");
            string godzRozp = comboBox1.Text.ToString();
            string dataZakon = dateTimePicker2.Value.ToString("dd-MM-yyyy");
            string godzZakon = comboBox4.Text.ToString();
            string min_roz = comboBox2.Text.ToString();
            string min_kon = comboBox3.Text.ToString();

            DateTime dat1 = dateTimePicker1.Value;
            DateTime dat2 = dateTimePicker2.Value;

            

            string razemDni;
            string razemGod;
            string razemMin;



            int dni = (int)((dat2 - dat1).TotalDays);
            int roz_godz = comboBox4.SelectedIndex - comboBox1.SelectedIndex;
            int roz_min = comboBox3.SelectedIndex - comboBox2.SelectedIndex;

       



            if ((dni < 0) || ((dni == 0) && (roz_godz == 0) && (roz_min <= 0)) || ((dni == 0) && (roz_godz < 0)))
            {
                MessageBox.Show("BŁĄD! DATY LUB GODZINY MUSZĄ BYĆ RÓŻNE!!!", "BŁĄD WPROWADZANIA DELEGACJI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;

            }

            else
            {


                if (roz_min < 0)
                {
                    roz_godz = roz_godz - 1;
                    roz_min = (6 + comboBox3.SelectedIndex - comboBox2.SelectedIndex);
                }

                if (roz_godz < 0)
                {
                    dni = dni - 1;
                    roz_godz = 24 + roz_godz;

                }



                if (roz_godz == 24 && roz_min == 0)
                {
                    dni = dni + 1;
                    roz_godz = 0;

                }

            }
            razemDni = dni.ToString();
            razemGod = roz_godz.ToString();
            razemMin = (roz_min * 10).ToString();






            string sni = numericUpDown1.Value.ToString();
            int sni1 = (int)numericUpDown1.Value;
            string obi = numericUpDown2.Value.ToString();
            int obi1 = (int)numericUpDown2.Value;
            string kol = numericUpDown3.Value.ToString();
            int kol11 = (int)numericUpDown3.Value;

            //obliczenie diety:

            //POSILKI


            double dieta_posilki = 0;


            dieta_posilki = 45 * (sni1 * 0.25 + obi1 * 0.5 + kol11 * 0.25);

        

            //dieta mniej niz 1 dzien

            if ((dni < 1) || ((dni == 1) && (roz_godz == 0) && (roz_min == 0)))
            {

                if ((roz_godz < 8) || (roz_godz == 8 && roz_min == 0))
                {

                    dieta = 0;
                }

                if (((roz_godz == 8) && (roz_min > 0)) || ((roz_godz > 8) && (roz_godz < 12)) || ((roz_godz == 12) && (roz_min == 0)))
                {

                    dieta = 22.5 - dieta_posilki;
                }

                if ((roz_godz > 12) || (roz_godz == 12 && roz_min > 0) || (roz_godz == 0))
                {

                    dieta = 45 - dieta_posilki;
                }



            }

            else

            //dieta wiecej niz 1 dzien
            {
                dieta = dni * 45;

                if ((roz_godz == 0) && (roz_min == 0))
                {
                    dieta = dni * 45-45;
                    
                }
                if (((roz_godz < 8) || (roz_godz == 8 && roz_min == 0)) || ((roz_godz == 0) && (roz_min != 0)))
                {

                    dieta = dieta + 22.5 - dieta_posilki;
                }

                else

                {

                    dieta = dieta + 45 - dieta_posilki;
                }


            }

            // zliczenie sumy z diety//



            double suma = double.Parse(label35.Text.ToString());

            if(dieta<0)
            {
                dieta = 0;
            }
            suma = suma + dieta;

            label35.Text = suma.ToString("N2");

            string y = dieta.ToString("N2");

            string ia = comboBox20.SelectedItem.ToString();


            string[] row = new string[] { "X",dataRozp, godzRozp + ":" + min_roz, dataZakon, godzZakon + ":" + min_kon, razemDni + " dni " + razemGod + " h " + razemMin + "m", textBox17.Text, sni, obi, kol, y, "66640", ia.Substring(0,5) };
            delegacjeKrajowa.Rows.Add(row);
       
        }

 

        // delegacje ZAGRANICZNE

        //NOCLEGI
        private void button4_Click(object sender, EventArgs e)
        {
            if (comboBox6.SelectedIndex==0)
            {
                return;
            }

        
            double sumaPln = 0;

            double ilosc = double.Parse(textBox7.Text.ToString());
            double suma= double.Parse(textBox8.Text.ToString());
            double kurs= double.Parse(textBox9.Text.ToString());
            string ia = comboBox12.SelectedItem.ToString();


            if (ilosc<0)
            {
                ilosc = (-1) * ilosc;
                textBox7.Text = ilosc.ToString();
            }

            if (suma < 0)
            {
                suma = (-1) * suma;
                textBox8.Text = suma.ToString();
            }

            if(ilosc==0 || suma==0)
            {
                return;
            }
            sumaPln = suma * kurs;


            if (radioButton1.Checked)
            {
                string[] row = new string[] { "", "NOCLEG", textBox7.Text, suma.ToString("N2"), comboBox6.SelectedItem.ToString(), textBox9.Text, sumaPln.ToString("N2"), "66670", ia.Substring(0, 5) };

                noclegiZagraniczna.Rows.Add(row);
            }
            else
            {

                double ryczalt = 0;
                int ilosc1 = int.Parse(textBox7.Text);
                double wart1 = double.Parse(textBox8.Text);

                ryczalt = ilosc1 * wart1;

                string[] row = new string[] { "", "RYCZAŁT", textBox7.Text, ryczalt.ToString("N2"), comboBox6.SelectedItem.ToString(), textBox9.Text, (kurs*ryczalt).ToString("N2"), "66640", ia.Substring(0, 5) };

                noclegiZagraniczna.Rows.Add(row);
            }

           label27.Text = sumuj(noclegiZagraniczna, 6).ToString("N2");
           label25.Text = sumuj(noclegiZagraniczna, 2).ToString();
        }

   

        private void textBox8_Validated(object sender, EventArgs e)
        {
          


            if (czyNumeric(textBox8.Text, "d"))
            {

                double s = double.Parse(textBox8.Text.ToString());
                textBox8.Text = s.ToString("N2");
            }

            else
            {

                textBox8.Text = "0,00";
            }

      

        }

        private void textBox7_Validated(object sender, EventArgs e)
        {
            if (czyNumeric(textBox7.Text, "n"))
            {
                int s = int.Parse(textBox7.Text.ToString());

                textBox7.Text = s.ToString();

            }

            else
            {

                textBox7.Text = "0";
            }
        }


        //DODATKOWE
        private void button5_Click(object sender, EventArgs e)
        {
            

            if (comboBox8.SelectedIndex==0 || comboBox7.SelectedIndex == 0)
            {
                return;
            }
            
            double sumaPln = 0;

   
            double suma = double.Parse(textBox11.Text.ToString());
            if( suma<=0)
            {
                return;
            }
            double kurs = double.Parse(textBox10.Text.ToString());
            string konto= kontoDodatkowezagraniczne(comboBox8.SelectedIndex);
            string ia = comboBox17.SelectedItem.ToString();

            sumaPln = suma * kurs;
            string[] row = new string[] { "", comboBox8.SelectedItem.ToString(), suma.ToString("N2"), comboBox7.SelectedItem.ToString(), textBox10.Text, sumaPln.ToString("N2"), konto,ia.Substring(0,5)};
               dodatkoweZagraniczna.Rows.Add(row);

            label44.Text = sumuj(dodatkoweZagraniczna, 5).ToString("N2");
     
        }

        private void textBox11_Validated(object sender, EventArgs e)
        {

            if (czyNumeric(textBox11.Text, "d"))
            {

                double s = double.Parse(textBox11.Text.ToString());
                textBox11.Text = s.ToString("N2");
            }

            else
            {

                textBox11.Text = "0,00";
            }

        }

       
       string kontoDodatkowezagraniczne(int index)
       {
           string konto="";


           switch(index)
           {
               case 1:
                   konto = "68150";
                break;
                
               case 2:
               case 3:
               case 10:
                    konto = "66418";
                break;
               case 4:
               case 5:
               case 6:
               case 7:
               case 8:
                konto = "66630";
                break;
               case 9:
                konto = "65500";
                break;
                case 11:
                    konto = "66400";
                    break;
                case 12:
                    konto = "66418";
                    break;
                case 13:
                    konto = "68160";
                    break;

            }
            //  -- WYBIERZ KOSZT --
            //Artykuły biurowe za granicą	68150
            //Rozmowy handlowe za granicą - bowling	66418
            //Rozmowy handlowe za granicą - rachunki z restauracji	66418
            //Opłata parkingowa za granicą	66630
            //Pozostałe koszty dodatkowe podróży za granicą	66630
            //Koszty taksówek za granicą	66630
            //Bilet wstępu na targi za granicą	66630
            //Prom, opłata drogowa, pociąg, samolot, autostrada, vinieta za granicą	66630
            //Wynajem samochodu za granicą	65500
            //Napiwki 	66418
            return konto;
       }


     // samochodu

       private void button6_Click(object sender, EventArgs e)
       {
           if (comboBox9.SelectedIndex == 0 || comboBox10.SelectedIndex == 0)
           {
               return;
           }

           double sumaPln = 0;


           double suma = double.Parse(textBox13.Text.ToString());
            if(suma<=0)
            {
                return;
            }
           double kurs = double.Parse(textBox12.Text.ToString());
           string konto = kontoSamochoduzagraniczne(comboBox9.SelectedIndex);
           string ia = comboBox18.SelectedItem.ToString();

           sumaPln = suma * kurs;
           string[] row = new string[] { "", comboBox9.SelectedItem.ToString(), suma.ToString("N2"), comboBox10.SelectedItem.ToString(), textBox12.Text, sumaPln.ToString("N2"), konto, ia.Substring(0,5) };
           samochoduZagraniczna.Rows.Add(row);

           label1.Text = sumuj(samochoduZagraniczna, 5).ToString("N2");
       }

       private void textBox13_Validated(object sender, EventArgs e)
       {

           if (czyNumeric(textBox13.Text, "d"))
           {

               double s = double.Parse(textBox13.Text.ToString());
               textBox13.Text = s.ToString("N2");
           }

           else
           {

               textBox13.Text = "0,00";
           }
       }

       string kontoSamochoduzagraniczne(int index)
       {
           string konto = "";


           switch (index)
           {
               case 1:
                   konto = "65100";
                   break;

               case 2:
               case 3:
                   konto = "65000";
                   break;
           }


           return konto;
       }

        // DATA GRANICZNA KURSU 
    




       private void comboBox6_SelectedValueChanged(object sender, EventArgs e)
       {
           if (comboBox6.SelectedIndex != 0)
           {
               textBox9.Text = szukajKursu(comboBox6.SelectedItem.ToString());
           }
           }

       private void comboBox7_SelectedValueChanged(object sender, EventArgs e)
       {
           if (comboBox7.SelectedIndex != 0)
           {
               textBox10.Text = szukajKursu(comboBox7.SelectedItem.ToString());
           }
       }

       private void comboBox10_SelectedIndexChanged(object sender, EventArgs e)
       {
           if (comboBox10.SelectedIndex != 0)
           {
               textBox12.Text = szukajKursu(comboBox10.SelectedItem.ToString());
           }
       }





       string pobierzDaneNBP(string data)
       {


       
           if (!File.Exists("dir.txt")) 
           {

            
               System.Net.WebClient wb = new System.Net.WebClient();
               wb.DownloadFile("http://www.nbp.pl/Kursy/xml/dir.txt", "dir.txt");
           }
          


           StreamReader objReader = new StreamReader("dir.txt");

           String sLine = "";
           string nazwaPliku ="";


           while (sLine != null)
           {
               sLine = objReader.ReadLine();
               if (sLine != null)
               {

                   if (sLine.StartsWith("a") && sLine.EndsWith(data)) 
                   
                   {

                       nazwaPliku = sLine.ToString();

                       MessageBox.Show(nazwaPliku.ToString());
                       return nazwaPliku;


                   }

               }

           }
           objReader.Close();


          

           return "";



       }



 
       private void button7_Click(object sender, EventArgs e)
       {
            

            if ( comboBox11.SelectedIndex == 0)
           {
               return;
           }


            if (textBox15.Text.Length<6)
            {

                MessageBox.Show("Wprowadź cel delegacji. min 7 znaków ", "CEL DELEGACJI",    MessageBoxButtons.OK, MessageBoxIcon.Error);
                
                return;
            }

            //wartosc diety
            double dieta = 0;


            //wprowadzenie rekordu delegacji


            string dataRozp = dateTimePicker4.Value.ToString("dd-MM-yyyy");
            string godzRozp = comboBox13.Text.ToString();
            string dataZakon = dateTimePicker5.Value.ToString("dd-MM-yyyy");
            string godzZakon = comboBox15.Text.ToString();
            string min_roz = comboBox14.Text.ToString();
            string min_kon = comboBox16.Text.ToString();

            DateTime dat3 = dateTimePicker4.Value;
            DateTime dat4 = dateTimePicker5.Value;


   

            string razemDni;
            string razemGod;
            string razemMin;

            //nowowsc rozwiazuje problem niezaznaczenie dtp4

            //TimeSpan ts = dat4-dat3;
            //int dni = int.Parse(ts.Days.ToString());

            //tak było wcześcniej
            int dni = (int)((dat4 - dat3).TotalDays);
            int roz_godz = comboBox15.SelectedIndex - comboBox13.SelectedIndex;
            int roz_min = comboBox16.SelectedIndex - comboBox14.SelectedIndex;


          


            if ((dni < 0) || ((dni == 0) && (roz_godz == 0) && (roz_min <= 0)) || ((dni == 0) && (roz_godz < 0)))
            {
                MessageBox.Show("BŁĄD! DATY LUB GODZINY MUSZĄ BYĆ RÓŻNE!!!", "BŁĄD WPROWADZANIA DELEGACJI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;

            }

         
                 else
            {


                if (roz_min < 0)
                {
                    roz_godz = roz_godz - 1;
                    roz_min = (6 + comboBox16.SelectedIndex - comboBox14.SelectedIndex);
                }

                if (roz_godz < 0)
                {
                    dni = dni - 1;
                    roz_godz = 24 + roz_godz;

                }



                if (roz_godz == 24 && roz_min == 0)
                {
                    dni = dni + 1;
                    roz_godz = 0;

                }

            }
            razemDni = dni.ToString();
            razemGod = roz_godz.ToString();
            razemMin = (roz_min * 10).ToString();






            string sni = numericUpDown4.Value.ToString();
            int sni1 = (int)numericUpDown4.Value;
            string obi = numericUpDown5.Value.ToString();
            int obi1 = (int)numericUpDown5.Value;
            string kol = numericUpDown6.Value.ToString();
            int kol11 = (int)numericUpDown6.Value;
            string kraj = comboBox11.SelectedItem.ToString();
            double kurs = double.Parse(textBox14.Text);


           

            //obliczenie diety:

            //POSILKI


            double dieta_posilki = 0;


            dieta_posilki = szukajdiety(kraj, 3) * (sni1 * 0.15 + obi1 * 0.3 + kol11 * 0.3);

            //dieta mniej niz 1 dzien
            if ((dni < 1) )
            {

                if ((roz_godz < 8) || (roz_godz == 8 && roz_min == 0))
                {

                    dieta = szukajdiety(kraj, 1) - dieta_posilki;
                }

                if (((roz_godz == 8) && (roz_min > 0)) || ((roz_godz > 8) && (roz_godz < 12)) || ((roz_godz == 12) && (roz_min == 0)))
                {

                    dieta = szukajdiety(kraj, 2) - dieta_posilki;
                }

                if ((roz_godz > 12) || (roz_godz == 12 && roz_min > 0) )
                {

                    dieta = szukajdiety(kraj, 3) - dieta_posilki;
                }

            }

                //dieta wiecej niz 1 dzien
            else if  ( (roz_godz == 0) && (roz_min == 0))
            {

               dieta = dni*szukajdiety(kraj, 3) - dieta_posilki; 
            }
            


            else 

             {
                if (((roz_godz > 0)&&(roz_godz < 8)) || (roz_godz == 8 && roz_min == 0) || (roz_godz == 0) && (roz_min != 0))
                {

                    dieta = dni * szukajdiety(kraj, 3) +szukajdiety(kraj, 1) - dieta_posilki;
                }

                if (((roz_godz == 8) && (roz_min > 0)) || ((roz_godz > 8) && (roz_godz < 12)) || ((roz_godz == 12) && (roz_min == 0)))
                {


                    dieta = dni * szukajdiety(kraj, 3) +szukajdiety(kraj, 2) - dieta_posilki;
             
                }

                if ((roz_godz > 12) || (roz_godz == 12 && roz_min > 0))
                {

                    dieta = dni * szukajdiety(kraj, 3) +szukajdiety(kraj, 3) - dieta_posilki;
                }

       
            }

            // zliczenie sumy z diety//

           if (dieta<0)
           {
               dieta = 0;
           }

            double suma = double.Parse(label52.Text.ToString());
            suma = suma + dieta;

            label52.Text = suma.ToString("N2");

            string y = (dieta*kurs).ToString("N2");
            string dw = dieta.ToString("N2");

            string ia = comboBox19.SelectedItem.ToString();

            string[] row = new string[] { "X", dataRozp, godzRozp + ":" + min_roz, dataZakon, godzZakon + ":" + min_kon, razemDni + " dni " + razemGod + " h " + razemMin + "m", textBox15.Text, comboBox11.SelectedItem.ToString(),textBox16.Text,textBox14.Text, sni, obi, kol, dw,y, "66640",ia.Substring(0,5) };
            delegacjeZagraniczna.Rows.Add(row);
            
    

           label52.Text = sumuj(delegacjeZagraniczna, 14).ToString("N2");
       }

       private void comboBox13_SelectedIndexChanged(object sender, EventArgs e)
       {

       }

       private void comboBox11_SelectedIndexChanged(object sender, EventArgs e)
       {
           if (comboBox11.SelectedIndex != 0)
           {
               textBox16.Text = szukajwaluty(comboBox11.SelectedItem.ToString());
               textBox14.Text = szukajKursu(textBox16.Text);

           }


       }

      
    
        string  szukajwaluty(string kraj)
       {
           string waluta = "";
      



           foreach (DataGridViewRow row in stawkiZagraniczna.Rows)
           {
        

               if (row.Cells[0].Value.ToString().Equals(kraj))
               {
               
                   waluta = row.Cells[1].Value.ToString();
                    
                   return waluta;
               }
           }

           waluta = "";
           return waluta;

       }

    
        double szukajdiety(string kraj, int stawka)
        {
            double dieta = 0;

            foreach (DataGridViewRow row in stawkiZagraniczna.Rows)
            {


                if (row.Cells[0].Value.ToString().Equals(kraj))
                {

                    dieta = double.Parse(row.Cells[stawka+1].Value.ToString());

                    return dieta;
                }
            }


            return dieta;

        }

        private void noclegiZagraniczna_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void textBox15_MouseClick(object sender, MouseEventArgs e)
        {
            ToolTip tt = new ToolTip();
     
            tt.Show("min. 7 znaków",textBox15,400);
        }

     

        private void textBox17_MouseClick(object sender, MouseEventArgs e)
        {
            ToolTip tt = new ToolTip();

            tt.Show("min. 7 znaków", textBox17, 400);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();


            form2.ShowDialog(this);

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            MessageBox.Show( Form2.data);
        }

        private void noclegiKrajowa_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if(radioButton2.Checked)
            {
                label62.Visible = true;
                comboBox5.Visible = true;
                textBox8.ReadOnly = true;
                comboBox6.Enabled = false;

            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                label62.Visible = false;
                comboBox5.Visible = false;
                textBox8.ReadOnly = false;
                comboBox6.Enabled = true;

            }
        }



        private void button1_Click_2(object sender, EventArgs e)
        {
            Form3 wczytywanie = new Form3();
            wczytywanie.Show();

        }

        private void comboBox5_SelectedValueChanged(object sender, EventArgs e)
        {

            if (comboBox5.SelectedIndex == 0)
            {

                return;
            }
       
      
            szukajRyczaltu(comboBox5.SelectedItem.ToString());
        }

        private void comboBox5_Leave(object sender, EventArgs e)
        {
            if (comboBox5.SelectedIndex == 0)
            {

                return;
            }


            szukajRyczaltu(comboBox5.SelectedItem.ToString());
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox5.SelectedIndex == 0)
            {

                return;
            }


            szukajRyczaltu(comboBox5.SelectedItem.ToString());
        }

        private void progressBar2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_3(object sender, EventArgs e)
        {
            if ( delegacjeKrajowa.Visible || noclegiKrajowa.Visible|| dodatkoweKrajowa.Visible|| samochoduKrajowa.Visible)
            {
                noclegiKrajowa.Rows.Clear();
                label17.Text = sumuj(noclegiKrajowa, 2).ToString();
                label18.Text = sumuj(noclegiKrajowa, 3).ToString("N2");
                dodatkoweKrajowa.Rows.Clear();
                label20.Text = sumuj(dodatkoweKrajowa, 2).ToString("N2");
                samochoduKrajowa.Rows.Clear();
                label22.Text = sumuj(samochoduKrajowa, 2).ToString("N2");
                 delegacjeKrajowa.Rows.Clear();
                label35.Text = sumuj(delegacjeKrajowa, 10).ToString("N2");

            }





            else
            {
                noclegiZagraniczna.Rows.Clear();
                label25.Text = "0";
                label27.Text = "0,00";

                dodatkoweZagraniczna.Rows.Clear();
                delegacjeZagraniczna.Rows.Clear();
                samochoduZagraniczna.Rows.Clear();

                label1.Text = "0,00";
                label52.Text = "0,00";
                label44.Text = "0,00";

            }
        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void mENUToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mojeDane.Visible = false;

        }

        private void label67_Click(object sender, EventArgs e)
        {

        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {


















































             
        }
    }
 
    }
        




      
      
  

