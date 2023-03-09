using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Xml;


namespace ROZLICZENIE_WERSJA_OSTATECZNA
{
    public partial class Form2 : Form
    {
       
        
        public static string zalKsiegowosc;
        public static string data;
        public static string data1;
        public static bool kursNapodstawiezaliczki = false;
        public static string nazwaTabeli;
        // zmienna, ktora wskazuje,czy nie bylo jeszcze uruchomionej dele zagranicznej
        public static bool pierwszeWywolanie= false;
   




        public Form2()
        {
            InitializeComponent();

            button3.Visible = false;

       

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            button3.Visible = false;
            ToolTip buttonToolTip = new ToolTip();

            buttonToolTip.ToolTipTitle = "WYBIERZ DATĘ";

            buttonToolTip.UseFading = true;

            buttonToolTip.UseAnimation = true;

            buttonToolTip.IsBalloon = true;



            buttonToolTip.ShowAlways = true;



            buttonToolTip.AutoPopDelay = 4000;

            buttonToolTip.InitialDelay = 800;

            buttonToolTip.ReshowDelay = 400;



            buttonToolTip.SetToolTip(dateTimePicker3, "Należy wybrać datę, ay kontynuować.");

            if (radioButton1.Checked == true)
            {
                dateTimePicker1.Visible = false;
                dateTimePicker3.Visible = true;
                label1.Visible = false;
                checkedListBox1.Visible = false;
                kursNapodstawiezaliczki = false;
                data1="";
                data = "";
                nazwaTabeli = "";

           


            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
          
            button3.Visible = false;
            ToolTip buttonToolTip = new ToolTip();

            buttonToolTip.ToolTipTitle = "WYBIERZ DATĘ";

            buttonToolTip.UseFading = true;

            buttonToolTip.UseAnimation = true;

            buttonToolTip.IsBalloon = true;



            buttonToolTip.ShowAlways = true;



            buttonToolTip.AutoPopDelay = 4000;

            buttonToolTip.InitialDelay = 800;

            buttonToolTip.ReshowDelay = 400;



            buttonToolTip.SetToolTip(dateTimePicker1, "Należy wybrać datę, ay kontynuować.");
            if (radioButton2.Checked == true)
            {
                dateTimePicker3.Visible = false;
                dateTimePicker1.Visible = true;
                label1.Visible = true;
                checkedListBox1.Visible = true;
                kursNapodstawiezaliczki = true;
                data1 = "";
                data = "";
                nazwaTabeli = "";
                
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

       
            this.Close();

        }

        private void button3_Click(object sender, EventArgs e)
        {

            if (!radioButton1.Checked&&!radioButton2.Checked)
            {
                MessageBox.Show("Wybierz sposób rozliczenia delegacji!", "BŁAD");
                return;
            }

          

            foreach (object itemChecked in checkedListBox1.CheckedItems)
            {
                zalKsiegowosc = zalKsiegowosc + ", " + itemChecked.ToString();

            }

            if (radioButton1.Checked && !radioButton2.Checked)
             {
                 
                stworzNazweTabeliNBP(dateTimePicker3);
                pierwszeWywolanie = true;
                this.Close();
            }


            if (radioButton2.Checked && !radioButton1.Checked)
             {
                 stworzNazweTabeliNBP(dateTimePicker1);
                pierwszeWywolanie = true;
                this.Close();
            }
          
            
        }

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {


            //button3.Visible = true;
            //stworzNazweTabeliNBP(dateTimePicker3);
           // data1 = dateTimePicker3.Value.Day.ToString() + "-" + dateTimePicker3.Value.Month.ToString() + "-" + dateTimePicker3.Value.Year.ToString();

           // MessageBox.Show("Data rozliczenia: " + data1 + Environment.NewLine+ " Data obowiązywania kursu z dnia: "+data + "");

        }



        void stworzNazweTabeliNBP (DateTimePicker dtp)
        {


            string rok = dtp.Value.Year.ToString();
            rok = rok.Substring(2, 2);
            string miesiac = dtp.Value.Month.ToString();
            int miesiacliczba = int.Parse(miesiac);
            if (miesiacliczba < 10)
            {
                miesiac = "0" + miesiac;
            }



            string dzien = dtp.Value.Day.ToString();
            int dzienliczba = int.Parse(dzien);


            if (dzienliczba < 10)
            {
                dzien = "0" + dzien;
            }

            // szukanie nazwy pliku

 



            string dataKursu = dataMinusdzien(rok + miesiac + dzien);
            string nazwa = szukajTabeliNBP(dataKursu);


           while( nazwa=="")
           {
               dataKursu = dataMinusdzien(dataKursu); ;
               nazwa = szukajTabeliNBP(dataKursu);
           }


            if (dataKursu != "")
            {
                data = dataKursu.Substring(4, 2) + "-" + dataKursu.Substring(2, 2) + "-" + "20" + dataKursu.Substring(0, 2);
                nazwaTabeli = nazwa;
            }
            if (nazwa=="" || data =="")
            {

                MessageBox.Show("brak kursu z tego dnia");
                return;

            }

        
        }

  


        string szukajTabeliNBP(string data)
        {

            string tabela = "";
            
       
            string line;
          

          
                System.IO.StreamReader file = new System.IO.StreamReader("dir.txt");
                 while ((line = file.ReadLine()) != null)
                 {
                     if (line.StartsWith("a") && line.EndsWith(data))
                     {

                         tabela = line.ToString();
                         file.Close();
                         return tabela;
                     }

              


                 }
           



            tabela = "";
            return tabela;









        }

        string dataMinusdzien(string data)
        {

            string nowadata="";

            string rok = DateTime.Today.Year.ToString();

            string miesiac = data.Substring(2,2);
            string dzien = data.Substring(4,2);

            DateTime dt1 = DateTime.Parse(rok+"-"+miesiac+"-"+dzien);

            int dzienroku = (int)dt1.DayOfYear;

            if(dzienroku==1)
            {
                return "";
            }
            dzienroku --;

            var dt = new DateTime(DateTime.Now.Year, 1, 1).AddDays(dzienroku - 1);

            rok = dt.Year.ToString();
            rok = rok.Substring(2, 2);
            miesiac = dt.Month.ToString();
            int miesiacliczba = int.Parse(miesiac);
            if (miesiacliczba < 10)
            {
                miesiac = "0" + miesiac;
            }



            dzien = dt.Day.ToString();
            int dzienliczba = int.Parse(dzien);


            if (dzienliczba < 10)
            {
                dzien = "0" + dzien;
            }

            nowadata = rok + miesiac + dzien;
            return nowadata;
 
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

           
         //   stworzNazweTabeliNBP(dateTimePicker1);
         //   string data1 = dateTimePicker1.Value.Day.ToString() + "-" + dateTimePicker1.Value.Month.ToString() + "-" + dateTimePicker1.Value.Year.ToString();

         //   MessageBox.Show("Data rozliczenia: " + data1 + Environment.NewLine + "Data obowiązywania kursu z dnia: " + data + "");
        //    button3.Visible = true;


        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

     

        private void dateTimePicker1_CloseUp(object sender, EventArgs e)
        {
            stworzNazweTabeliNBP(dateTimePicker1);
            string data1 = dateTimePicker1.Value.Day.ToString() + "-" + dateTimePicker1.Value.Month.ToString() + "-" + dateTimePicker1.Value.Year.ToString();

            MessageBox.Show("Data rozliczenia: " + data1 + Environment.NewLine + "Data obowiązywania kursu z dnia: " + data + "");
            button3.Visible = true;
        }

        private void dateTimePicker3_CloseUp(object sender, EventArgs e)
        {
            button3.Visible = true;
            stworzNazweTabeliNBP(dateTimePicker3);
             data1 = dateTimePicker3.Value.Day.ToString() + "-" + dateTimePicker3.Value.Month.ToString() + "-" + dateTimePicker3.Value.Year.ToString();

             MessageBox.Show("Data rozliczenia: " + data1 + Environment.NewLine+ " Data obowiązywania kursu z dnia: "+data + "");

        }
    }
}





