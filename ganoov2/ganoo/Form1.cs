using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace ganoo
{
    public partial class Form1 : Form
    {
        OleDbConnection baglanti = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb");
        OleDbCommand komut = new OleDbCommand();
        OleDbDataAdapter adtr = new OleDbDataAdapter();
        DataSet ds = new DataSet();

        public Form1()
        {
            InitializeComponent();
        }

        private string HarfNotuHesapla(double ort)
        {
            string harfNotu = "";
            if (ort <= 39 && ort >= 0) { harfNotu = "FF"; }
            else if (ort <= 44 && ort >= 40) { harfNotu = "FD"; }
            else if (ort <= 49 && ort >= 45) { harfNotu = "DD"; }
            else if (ort <= 59 && ort >= 50) { harfNotu = "DC"; }
            else if (ort <= 64 && ort >= 60) { harfNotu = "CC"; }
            else if (ort <= 74 && ort >= 65) { harfNotu = "CB"; }
            else if (ort <= 84 && ort >= 75) { harfNotu = "BB"; }
            else if (ort <= 89 && ort >= 85) { harfNotu = "BA"; }
            else if (ort <= 100 && ort >= 90) { harfNotu = "AA"; }

            return harfNotu;
        }

        private float AnoHesapla(int krediNotu, string harfNotu)
        {
            float ano = 0;
            switch (harfNotu)
            {
                case "AA": ano = krediNotu * 4f; break;
                case "BA": ano = krediNotu * 3.5f; break;
                case "BB": ano = krediNotu * 3f; break;
                case "CB": ano = krediNotu * 2.5f; break;
                case "CC": ano = krediNotu * 2f; break;
                case "DC": ano = krediNotu * 1.5f; break;
                case "DD": ano = krediNotu * 1f; break;
                case "FD": ano = krediNotu * 0.5f; break;
                case "FF": ano = krediNotu * 0f; break;
            }
            return ano;
        }

        private void Hesapla()
        {
            float donem1Ano = 0, donem2Ano = 0, donem3Ano = 0, donem4Ano = 0;

            for (int i = 1; i < 22; i++)
            {
                string vize = "vv" + i.ToString();
                string final = "ff" + i.ToString();
                string harfNotu = "harf" + i.ToString();
                string ortalamaLbl = "ort" + i.ToString();
                string kredi = "k" + i.ToString();
                if (this.Controls.Find(vize, false).First().Text == "" || this.Controls.Find(final, false).First().Text == "")
                {
                    continue;
                }

                int vizeNotu = Convert.ToInt32(this.Controls.Find(vize, false).First().Text);
                int finalNotu = Convert.ToInt32(this.Controls.Find(final, false).First().Text);
                double ortalama = vizeNotu * 0.4 + finalNotu * 0.6;

                string h_not = this.Controls.Find(harfNotu, false).First().Text;
                if (h_not == "" || buttonClick)
                {
                    h_not = HarfNotuHesapla(ortalama);
                    this.Controls.Find(harfNotu, false).First().Text = h_not;
                    this.Controls.Find(ortalamaLbl, false).First().Text = ortalama.ToString();
                }
                if (h_not == "AA" || h_not == "BA" || h_not == "BB")
                {
                    this.Controls.Find(harfNotu, false).First().ForeColor = Color.Green;
                }
                else if (h_not == "FF" || h_not == "FD")
                {
                    this.Controls.Find(harfNotu, false).First().ForeColor = Color.Red;
                }
                else if (h_not == "DD" || h_not == "DC")
                {
                    this.Controls.Find(harfNotu, false).First().ForeColor = Color.Yellow;
                }
                else if (h_not == "CC" || h_not == "CB")
                {
                    this.Controls.Find(harfNotu, false).First().ForeColor = Color.Orange;
                }
                int krediNotu = Convert.ToInt32(this.Controls.Find(kredi, false).First().Text);

                float ano = AnoHesapla(krediNotu, h_not);

                if (i <= 6) { donem1Ano += ano; }
                else if (i <= 11) { donem2Ano += ano; }
                else if (i <= 16) { donem3Ano += ano; }
                else if (i <= 21) { donem4Ano += ano; }
            }

            Donem1Ano.Text = (donem1Ano / 17).ToString();
            Donem2Ano.Text = (donem2Ano / 13).ToString();
            Donem3Ano.Text = (donem3Ano / 13).ToString();
            Donem4Ano.Text = (donem4Ano / 13).ToString();
            float toplamDonemAno = 0;
            int toplamKrediNotu = 0;
            if (donem1Ano > 0)
            {
                toplamDonemAno += donem1Ano;
                toplamKrediNotu += 17;
            }
            if (donem2Ano > 0)
            {
                toplamDonemAno += donem2Ano;
                toplamKrediNotu += 13;
            }
            if (donem3Ano > 0)
            {
                toplamDonemAno += donem3Ano;
                toplamKrediNotu += 13;
            }
            if (donem4Ano > 0)
            {
                toplamDonemAno += donem4Ano;
                toplamKrediNotu += 13;
            }
            float ganoNotu = toplamDonemAno / toplamKrediNotu;
            gano.Text = ganoNotu.ToString();
        }

        private bool buttonClick = false;
        private void button1_Click(object sender, EventArgs e)
        {
            buttonClick = true;
            Hesapla();
            buttonClick = false;
        }

        private void h10_TextChanged(object sender, EventArgs e)
        {
            if (buttonClick == false)
            {
                Hesapla();
                int satir = Convert.ToInt32(((Control)sender).Tag);
                string harfNotu = ((Control)sender).Text;
                if (harfNotu == "AA")
                {
                    this.Controls.Find("ort" + satir, false).First().Text = "90";
                }
                else if (harfNotu == "BA")
                {
                    this.Controls.Find("ort" + satir, false).First().Text = "85";
                }
                else if (harfNotu == "BB")
                {
                    this.Controls.Find("ort" + satir, false).First().Text = "75";
                }
                else if (harfNotu == "CB")
                {
                    this.Controls.Find("ort" + satir, false).First().Text = "65";
                }
                else if (harfNotu == "CC")
                {
                    this.Controls.Find("ort" + satir, false).First().Text = "60";
                }
                else if (harfNotu == "DC")
                {
                    this.Controls.Find("ort" + satir, false).First().Text = "50";
                }
                else if (harfNotu == "DD")
                {
                    this.Controls.Find("ort" + satir, false).First().Text = "45";
                }
                else if (harfNotu == "FD")
                {
                    this.Controls.Find("ort" + satir, false).First().Text = "40";
                }
                else if (harfNotu == "FF")
                {
                    this.Controls.Find("ort" + satir, false).First().Text = "0";
                }
                if (harf1.SelectedIndex == 0 || harf1.SelectedIndex == 1 || harf1.SelectedIndex == 2 || harf1.SelectedIndex == 3 || harf1.SelectedIndex == 4 || harf1.SelectedIndex == 5 || harf1.SelectedIndex == 6 || harf1.SelectedIndex == 7 || harf1.SelectedIndex == 8 || harf1.SelectedIndex == 9)
                {
                    richTextBox1.Text += d1.Text + " " + "dersinin notu " + " " + harf1.Text + " " + "iken" + " " + y1.Text + " " + "olarak" + " " + "değiştirilmiştir." + "\n";
                    harf1.Text = Convert.ToString(y1.Text);
                }
                if (harf2.SelectedIndex == 0 || harf2.SelectedIndex == 1 || harf2.SelectedIndex == 2 || harf2.SelectedIndex == 3 || harf2.SelectedIndex == 4 || harf2.SelectedIndex == 5 || harf2.SelectedIndex == 6 || harf2.SelectedIndex == 7 || harf2.SelectedIndex == 8 || harf2.SelectedIndex == 9)
                {
                    richTextBox1.Text += d2.Text + " " + "dersinin notu " + " " + harf2.Text + " " + "iken" + " " + y2.Text + " " + "olarak" + " " + "değiştirilmiştir." + "\n";
                    harf2.Text = Convert.ToString(y2.Text);
                }
                if (harf3.SelectedIndex == 0 || harf3.SelectedIndex == 1 || harf3.SelectedIndex == 2 || harf3.SelectedIndex == 3 || harf3.SelectedIndex == 4 || harf3.SelectedIndex == 5 || harf3.SelectedIndex == 6 || harf3.SelectedIndex == 7 || harf3.SelectedIndex == 8 || harf3.SelectedIndex == 9)
                {
                    richTextBox1.Text += d3.Text + " " + "dersinin notu " + " " + harf3.Text + " " + "iken" + " " + y3.Text + " " + "olarak" + " " + "değiştirilmiştir." + "\n";
                    harf3.Text = Convert.ToString(y3.Text);
                }
                if (harf4.SelectedIndex == 0 || harf4.SelectedIndex == 1 || harf4.SelectedIndex == 2 || harf4.SelectedIndex == 3 || harf4.SelectedIndex == 4 || harf4.SelectedIndex == 5 || harf4.SelectedIndex == 6 || harf4.SelectedIndex == 7 || harf4.SelectedIndex == 8 || harf4.SelectedIndex == 9)
                {
                    richTextBox1.Text += d4.Text + " " + "dersinin notu " + " " + harf4.Text + " " + "iken" + " " + y4.Text + " " + "olarak" + " " + "değiştirilmiştir." + "\n";
                    harf4.Text = Convert.ToString(y4.Text);
                }
                if (harf5.SelectedIndex == 0 || harf5.SelectedIndex == 1 || harf5.SelectedIndex == 2 || harf5.SelectedIndex == 3 || harf5.SelectedIndex == 4 || harf5.SelectedIndex == 5 || harf5.SelectedIndex == 6 || harf5.SelectedIndex == 7 || harf5.SelectedIndex == 8 || harf5.SelectedIndex == 9)
                {
                    richTextBox1.Text += d5.Text + " " + "dersinin notu " + " " + harf5.Text + " " + "iken" + " " + y5.Text + " " + "olarak" + " " + "değiştirilmiştir." + "\n";
                    harf5.Text = Convert.ToString(y5.Text);
                }
                if (harf6.SelectedIndex == 0 || harf6.SelectedIndex == 1 || harf6.SelectedIndex == 2 || harf6.SelectedIndex == 3 || harf6.SelectedIndex == 4 || harf6.SelectedIndex == 5 || harf6.SelectedIndex == 6 || harf6.SelectedIndex == 7 || harf6.SelectedIndex == 8 || harf6.SelectedIndex == 9)
                {
                    richTextBox1.Text += d6.Text + " " + "dersinin notu " + " " + harf6.Text + " " + "iken" + " " + y6.Text + " " + "olarak" + " " + "değiştirilmiştir." + "\n";
                    harf6.Text = Convert.ToString(y6.Text);
                }
                if (harf7.SelectedIndex == 0 || harf7.SelectedIndex == 1 || harf7.SelectedIndex == 2 || harf7.SelectedIndex == 3 || harf7.SelectedIndex == 4 || harf7.SelectedIndex == 5 || harf7.SelectedIndex == 6 || harf7.SelectedIndex == 7 || harf7.SelectedIndex == 8 || harf7.SelectedIndex == 9)
                {
                    richTextBox1.Text += d7.Text + " " + "dersinin notu " + " " + harf7.Text + " " + "iken" + " " + y7.Text + " " + "olarak" + " " + "değiştirilmiştir." + "\n";
                    harf7.Text = Convert.ToString(y7.Text);
                }
                if (harf8.SelectedIndex == 0 || harf8.SelectedIndex == 1 || harf8.SelectedIndex == 2 || harf8.SelectedIndex == 3 || harf8.SelectedIndex == 4 || harf8.SelectedIndex == 5 || harf8.SelectedIndex == 6 || harf8.SelectedIndex == 7 || harf8.SelectedIndex == 8 || harf8.SelectedIndex == 9)
                {
                    richTextBox1.Text += d8.Text + " " + "dersinin notu " + " " + harf8.Text + " " + "iken" + " " + y8.Text + " " + "olarak" + " " + "değiştirilmiştir." + "\n";
                    harf8.Text = Convert.ToString(y8.Text);
                }
                if (harf9.SelectedIndex == 0 || harf9.SelectedIndex == 1 || harf9.SelectedIndex == 2 || harf9.SelectedIndex == 3 || harf9.SelectedIndex == 4 || harf9.SelectedIndex == 5 || harf9.SelectedIndex == 6 || harf9.SelectedIndex == 7 || harf9.SelectedIndex == 8 || harf9.SelectedIndex == 9)
                {
                    richTextBox1.Text += d9.Text + " " + "dersinin notu " + " " + harf9.Text + " " + "iken" + " " + y9.Text + " " + "olarak" + " " + "değiştirilmiştir." + "\n";
                    harf9.Text = Convert.ToString(y9.Text);
                }
                if (harf10.SelectedIndex == 0 || harf10.SelectedIndex == 1 || harf10.SelectedIndex == 2 || harf10.SelectedIndex == 3 || harf10.SelectedIndex == 4 || harf10.SelectedIndex == 5 || harf10.SelectedIndex == 6 || harf10.SelectedIndex == 7 || harf10.SelectedIndex == 8 || harf10.SelectedIndex == 9)
                {
                    richTextBox1.Text += d10.Text + " " + "dersinin notu " + " " + harf10.Text + " " + "iken" + " " + y10.Text + " " + "olarak" + " " + "değiştirilmiştir." + "\n";
                    harf10.Text = Convert.ToString(y10.Text);
                }
                if (harf11.SelectedIndex == 0 || harf11.SelectedIndex == 1 || harf11.SelectedIndex == 2 || harf11.SelectedIndex == 3 || harf11.SelectedIndex == 4 || harf11.SelectedIndex == 5 || harf11.SelectedIndex == 6 || harf11.SelectedIndex == 7 || harf11.SelectedIndex == 8 || harf11.SelectedIndex == 9)
                {
                    richTextBox1.Text += d11.Text + " " + "dersinin notu " + " " + harf11.Text + " " + "iken" + " " + y11.Text + " " + "olarak" + " " + "değiştirilmiştir." + "\n";
                    harf11.Text = Convert.ToString(y11.Text);
                }
                if (harf12.SelectedIndex == 0 || harf12.SelectedIndex == 1 || harf12.SelectedIndex == 2 || harf12.SelectedIndex == 3 || harf12.SelectedIndex == 4 || harf12.SelectedIndex == 5 || harf12.SelectedIndex == 6 || harf12.SelectedIndex == 7 || harf12.SelectedIndex == 8 || harf12.SelectedIndex == 9)
                {
                    richTextBox1.Text += d12.Text + " " + "dersinin notu " + " " + harf12.Text + " " + "iken" + " " + y12.Text + " " + "olarak" + " " + "değiştirilmiştir." + "\n";
                    harf12.Text = Convert.ToString(y12.Text);
                }
                if (harf13.SelectedIndex == 0 || harf13.SelectedIndex == 1 || harf13.SelectedIndex == 2 || harf13.SelectedIndex == 3 || harf13.SelectedIndex == 4 || harf13.SelectedIndex == 5 || harf13.SelectedIndex == 6 || harf13.SelectedIndex == 7 || harf13.SelectedIndex == 8 || harf13.SelectedIndex == 9)
                {
                    richTextBox1.Text += d13.Text + " " + "dersinin notu " + " " + harf13.Text + " " + "iken" + " " + y13.Text + " " + "olarak" + " " + "değiştirilmiştir." + "\n";
                    harf13.Text = Convert.ToString(y13.Text);
                }
                if (harf14.SelectedIndex == 0 || harf14.SelectedIndex == 1 || harf14.SelectedIndex == 2 || harf14.SelectedIndex == 3 || harf14.SelectedIndex == 4 || harf14.SelectedIndex == 5 || harf14.SelectedIndex == 6 || harf14.SelectedIndex == 7 || harf14.SelectedIndex == 8 || harf14.SelectedIndex == 9)
                {
                    richTextBox1.Text += d14.Text + " " + "dersinin notu " + " " + harf14.Text + " " + "iken" + " " + y14.Text + " " + "olarak" + " " + "değiştirilmiştir." + "\n";
                    harf1.Text = Convert.ToString(y1.Text);
                }
                if (harf15.SelectedIndex == 0 || harf15.SelectedIndex == 1 || harf15.SelectedIndex == 2 || harf15.SelectedIndex == 3 || harf15.SelectedIndex == 4 || harf15.SelectedIndex == 5 || harf15.SelectedIndex == 6 || harf15.SelectedIndex == 7 || harf15.SelectedIndex == 8 || harf15.SelectedIndex == 9)
                {
                    richTextBox1.Text += d15.Text + " " + "dersinin notu " + " " + harf15.Text + " " + "iken" + " " + y15.Text + " " + "olarak" + " " + "değiştirilmiştir." + "\n";
                    harf15.Text = Convert.ToString(y15.Text);
                }
                if (harf16.SelectedIndex == 0 || harf16.SelectedIndex == 1 || harf16.SelectedIndex == 2 || harf16.SelectedIndex == 3 || harf16.SelectedIndex == 4 || harf16.SelectedIndex == 5 || harf16.SelectedIndex == 6 || harf16.SelectedIndex == 7 || harf16.SelectedIndex == 8 || harf16.SelectedIndex == 9)
                {
                    richTextBox1.Text += d16.Text + " " + "dersinin notu " + " " + harf16.Text + " " + "iken" + " " + y16.Text + " " + "olarak" + " " + "değiştirilmiştir." + "\n";
                    harf16.Text = Convert.ToString(y16.Text);
                }
                if (harf17.SelectedIndex == 0 || harf17.SelectedIndex == 1 || harf17.SelectedIndex == 2 || harf17.SelectedIndex == 3 || harf17.SelectedIndex == 4 || harf17.SelectedIndex == 5 || harf17.SelectedIndex == 6 || harf17.SelectedIndex == 7 || harf17.SelectedIndex == 8 || harf17.SelectedIndex == 9)
                {
                    richTextBox1.Text += d17.Text + " " + "dersinin notu " + " " + harf17.Text + " " + "iken" + " " + y17.Text + " " + "olarak" + " " + "değiştirilmiştir." + "\n";
                    harf17.Text = Convert.ToString(y17.Text);
                }
                if (harf18.SelectedIndex == 0 || harf18.SelectedIndex == 1 || harf18.SelectedIndex == 2 || harf18.SelectedIndex == 3 || harf18.SelectedIndex == 4 || harf18.SelectedIndex == 5 || harf18.SelectedIndex == 6 || harf18.SelectedIndex == 7 || harf18.SelectedIndex == 8 || harf18.SelectedIndex == 9)
                {
                    richTextBox1.Text += d18.Text + " " + "dersinin notu " + " " + harf18.Text + " " + "iken" + " " + y18.Text + " " + "olarak" + " " + "değiştirilmiştir." + "\n";
                    harf18.Text = Convert.ToString(y18.Text);
                }
                if (harf19.SelectedIndex == 0 || harf1.SelectedIndex == 1 || harf19.SelectedIndex == 2 || harf19.SelectedIndex == 3 || harf19.SelectedIndex == 4 || harf19.SelectedIndex == 5 || harf19.SelectedIndex == 6 || harf19.SelectedIndex == 7 || harf19.SelectedIndex == 8 || harf19.SelectedIndex == 9)
                {
                    richTextBox1.Text += d19.Text + " " + "dersinin notu " + " " + harf19.Text + " " + "iken" + " " + y19.Text + " " + "olarak" + " " + "değiştirilmiştir." + "\n";
                    harf19.Text = Convert.ToString(y19.Text);
                }
                if (harf20.SelectedIndex == 0 || harf20.SelectedIndex == 1 || harf20.SelectedIndex == 2 || harf20.SelectedIndex == 3 || harf20.SelectedIndex == 4 || harf20.SelectedIndex == 5 || harf20.SelectedIndex == 6 || harf20.SelectedIndex == 7 || harf20.SelectedIndex == 8 || harf20.SelectedIndex == 9)
                {
                    richTextBox1.Text += d20.Text + " " + "dersinin notu " + " " + harf20.Text + " " + "iken" + " " + y1.Text + " " + "olarak" + " " + "değiştirilmiştir." + "\n";
                    harf20.Text = Convert.ToString(y20.Text);
                }
                if (harf21.SelectedIndex == 0 || harf21.SelectedIndex == 1 || harf21.SelectedIndex == 2 || harf21.SelectedIndex == 3 || harf21.SelectedIndex == 4 || harf21.SelectedIndex == 5 || harf21.SelectedIndex == 6 || harf21.SelectedIndex == 7 || harf21.SelectedIndex == 8 || harf21.SelectedIndex == 9)
                {
                    richTextBox1.Text += d21.Text + " " + "dersinin notu " + " " + harf21.Text + " " + "iken" + " " + y21.Text + " " + "olarak" + " " + "değiştirilmiştir." + "\n";
                    harf21.Text = Convert.ToString(y21.Text);
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.ogrencilerTableAdapter.Fill(this.database1DataSet.ogrenciler);
            string[] notlar = { "AA", "BA", "BB", "CB", "CC", "DC", "DD", "FD", "FF" };
            harf7.Items.AddRange(notlar); harf8.Items.AddRange(notlar);
            harf9.Items.AddRange(notlar); harf10.Items.AddRange(notlar);
            harf11.Items.AddRange(notlar); harf17.Items.AddRange(notlar);
            harf18.Items.AddRange(notlar); harf19.Items.AddRange(notlar);
            harf20.Items.AddRange(notlar); harf21.Items.AddRange(notlar);
            harf1.Items.AddRange(notlar); harf2.Items.AddRange(notlar);
            harf3.Items.AddRange(notlar); harf4.Items.AddRange(notlar);
            harf5.Items.AddRange(notlar); harf6.Items.AddRange(notlar);
            harf12.Items.AddRange(notlar); harf13.Items.AddRange(notlar);
            harf14.Items.AddRange(notlar); harf15.Items.AddRange(notlar);
            harf16.Items.AddRange(notlar);
            y1.Items.AddRange(notlar); y2.Items.AddRange(notlar);
            y3.Items.AddRange(notlar); y4.Items.AddRange(notlar);
            y5.Items.AddRange(notlar); y6.Items.AddRange(notlar);
            y7.Items.AddRange(notlar); y8.Items.AddRange(notlar);
            y9.Items.AddRange(notlar); y10.Items.AddRange(notlar);
            y11.Items.AddRange(notlar); y12.Items.AddRange(notlar);
            y13.Items.AddRange(notlar); y14.Items.AddRange(notlar);
            y15.Items.AddRange(notlar); y16.Items.AddRange(notlar);
            y17.Items.AddRange(notlar); y18.Items.AddRange(notlar);
            y19.Items.AddRange(notlar); y20.Items.AddRange(notlar);
            y21.Items.AddRange(notlar);

            this.KeyPreview = true;
            this.KeyDown += Form1_KeyDown;

            void listele()
            {
                baglanti.Open();
                OleDbDataAdapter adtr = new OleDbDataAdapter("Select * From ogrenciler", baglanti);

                adtr.Fill(ds, "ogrenciler");
                dataGridView1.DataSource = ds.Tables["ogrenciler"];
                adtr.Dispose();
                baglanti.Close();
            }
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            { 
                {
                    Directory.CreateDirectory(@"c:\\Users\\xxx\\Desktop\\pdf");

                    BaseFont STF_Helvetica_Turkish = BaseFont.CreateFont("Helvetica", "CP1254", BaseFont.NOT_EMBEDDED);
                    iTextSharp.text.Font fontNormal = new iTextSharp.text.Font(STF_Helvetica_Turkish, 9, iTextSharp.text.Font.NORMAL);

                    Document document = new Document();
                    PdfWriter.GetInstance(document, new FileStream("c:\\Users\\xxx\\Desktop\\pdf\\Transkript.pdf", FileMode.Create));
                    document.Open();

                    PdfPTable table = new PdfPTable(6);
                    table.AddCell("Dersler");
                    table.AddCell("Vize");
                    table.AddCell("Final");
                    table.AddCell("Ortalama");
                    table.AddCell("Harf Notu");
                    table.AddCell("Kredi");

                    for (int i = 1; i <= 21; i++)
                    {
                        string vize = "vv" + i.ToString();
                        string final = "ff" + i.ToString();
                        string harfNotu = "harf" + i.ToString();
                        string ortalamaLbl = "ort" + i.ToString();
                        string kredi = "k" + i.ToString();
                        string ders = "d" + i.ToString();

                        string dersAdi = this.Controls.Find(ders, false).First().Text;
                        string vizeNotu = this.Controls.Find(vize, false).First().Text;
                        string finalNotu = this.Controls.Find(final, false).First().Text;
                        string ortalama = this.Controls.Find(ortalamaLbl, false).First().Text;
                        string harf = this.Controls.Find(harfNotu, false).First().Text;
                        string krediNotu = this.Controls.Find(kredi, false).First().Text;

                        table.AddCell(new Phrase(dersAdi, fontNormal));
                        table.AddCell(vizeNotu);
                        table.AddCell(finalNotu);
                        table.AddCell(ortalama);
                        table.AddCell(harf);
                        table.AddCell(krediNotu);

                        if (i == 6)
                        {
                            table.AddCell("ANO");
                            PdfPCell cell = new PdfPCell(new Phrase(Donem1Ano.Text));
                            cell.Colspan = 5;
                            cell.HorizontalAlignment = 1;
                            table.AddCell(cell);
                        }
                        if (i == 11)
                        {
                            table.AddCell("ANO");
                            PdfPCell cell = new PdfPCell(new Phrase(Donem2Ano.Text));
                            cell.Colspan = 5;
                            cell.HorizontalAlignment = 1;
                            table.AddCell(cell);
                        }
                        if (i == 16)
                        {
                            table.AddCell("ANO");
                            PdfPCell cell = new PdfPCell(new Phrase(Donem3Ano.Text));
                            cell.Colspan = 5;
                            cell.HorizontalAlignment = 1;
                            table.AddCell(cell);
                        }
                        if (i == 21)
                        {
                            table.AddCell("ANO");
                            PdfPCell cell = new PdfPCell(new Phrase(Donem4Ano.Text));
                            cell.Colspan = 5;
                            cell.HorizontalAlignment = 1;
                            table.AddCell(cell);
                        }
                    }
                    document.Add(table);
                    document.Close();
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {   
            {
                {
                    Directory.CreateDirectory(@"c:\\Users\\xxx\\Desktop\\Alınacak Dersler");

                    BaseFont STF_Helvetica_Turkish = BaseFont.CreateFont("Helvetica", "CP1254", BaseFont.NOT_EMBEDDED);

                    iTextSharp.text.Font fontNormal = new iTextSharp.text.Font(STF_Helvetica_Turkish, 9, iTextSharp.text.Font.NORMAL);

                    Document document = new Document();
                    PdfWriter.GetInstance(document, new FileStream("c:\\Users\\xxx\\Desktop\\Alınacak Dersler\\Transkript-kalan.pdf", FileMode.Create));
                    document.Open();

                    PdfPTable table = new PdfPTable(6);
                    table.AddCell("Dersler");
                    table.AddCell("Vize");
                    table.AddCell("Final");
                    table.AddCell("Ortalama");
                    table.AddCell("Harf Notu");
                    table.AddCell("Kredi");

                    for (int i = 1; i <= 21; i++)
                    {
                        string vize = "vv" + i.ToString();
                        string final = "ff" + i.ToString();
                        string harfNotu = "harf" + i.ToString();
                        string ortalamaLbl = "ort" + i.ToString();
                        string kredi = "k" + i.ToString();
                        string ders = "d" + i.ToString();

                        string dersAdi = this.Controls.Find(ders, false).First().Text;
                        string vizeNotu = this.Controls.Find(vize, false).First().Text;
                        string finalNotu = this.Controls.Find(final, false).First().Text;
                        string ortalama = this.Controls.Find(ortalamaLbl, false).First().Text;
                        string harf = this.Controls.Find(harfNotu, false).First().Text;
                        string krediNotu = this.Controls.Find(kredi, false).First().Text;

                        if (harf == "FF" || harf == "FD")
                        {
                            table.AddCell((new Phrase(dersAdi, fontNormal)));
                            table.AddCell(new Phrase(dersAdi, fontNormal));
                            table.AddCell(vizeNotu);
                            table.AddCell(finalNotu);
                            table.AddCell(ortalama);
                            table.AddCell(harf);
                            table.AddCell(krediNotu);
                        }
                    }
                    document.Add(table);
                    document.Close();
                }
            }
        }

        void listele()
        {
            baglanti.Open();
            OleDbDataAdapter adtr = new OleDbDataAdapter("Select * From ogrenciler", baglanti);

            adtr.Fill(ds, "ogrenciler");
            dataGridView1.DataSource = ds.Tables["ogrenciler"];
            adtr.Dispose();
            baglanti.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != string.Empty && textBox2.Text != string.Empty && textBox3.Text != string.Empty)
            {
                for (int i = 1; i <= 21; i++)
                {
                    string dersler = "d" + i.ToString();
                    string derslerr = this.Controls.Find(dersler, false).First().Text;
                    string harfler = "harf" + i.ToString();
                    string harflerr = this.Controls.Find(harfler, false).First().Text;
                    komut.Connection = baglanti;
                    komut.CommandText = "Insert Into Dersler (d_ad,ogr_no,harf_notu) Values ('" + derslerr + "','" + textBox3.Text + "','" + harflerr + "')";
                    baglanti.Open();
                    komut.ExecuteNonQuery();
                    komut.Dispose();
                    baglanti.Close();
                }
                if (textBox3.Text != "")
                {
                    komut.Connection = baglanti;
                    komut.CommandText = "Insert Into ogrenciler(ogr_no,gano,ogr_ad,ogr_soyad) Values ('" + textBox3.Text + "', '" + gano.Text + "','" + textBox1.Text + "','" + textBox2.Text + "')";
                    baglanti.Open();
                    komut.ExecuteNonQuery();
                    komut.Dispose();
                    baglanti.Close();
                    MessageBox.Show("Kayıt Başarılı");
                }
                else
                {
                    MessageBox.Show("Boş Alan Bırakmayınız");
                }
                ds.Clear();
                listele();
            }
            else
                MessageBox.Show("Boş Bırakılan Alan Mevcut Doldurunuz...(ad-soyad-öğrencino)");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (textBox3.Text != string.Empty && textBox1.Text != string.Empty && textBox2.Text != string.Empty)
            {
                baglanti.Open();
                komut.Connection = baglanti;
                komut.CommandText = "Update ogrenciler set ogr_ad='" + textBox1.Text + "',ogr_soyad='" + textBox2.Text + "',gano='" + gano.Text + "' where ogr_no=" + textBox3.Text + "";
                komut.ExecuteNonQuery();
                komut.Dispose();
                baglanti.Close();
            }
            else
                MessageBox.Show("Boş Alanlar Mevcut...(ögrencino-ögrenciad-ögrencisoyad)");
            ds.Clear();
            listele();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (textBox3.Text != string.Empty)
            {
                DialogResult c;
                c = MessageBox.Show("Silmek İstediğinizden Emin Misiniz ?", "Uyarı!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (c == DialogResult.Yes)
                {
                    baglanti.Open();
                    komut.Connection = baglanti;
                    komut.CommandText = "Delete From ogrenciler Where ogr_no=" + textBox3.Text + "";
                    komut.ExecuteNonQuery();
                    komut.Dispose();
                    baglanti.Close();

                    listele();
                    if (c == DialogResult.Yes)
                    {
                        baglanti.Open();
                        komut.Connection = baglanti;
                        komut.CommandText = "Delete From dersler Where ogr_no=" + textBox3.Text + "";
                        komut.ExecuteNonQuery();
                        komut.Dispose();
                        baglanti.Close();
                        ds.Tables["ogrenciler"].Clear();
                        listele();
                    }
                }
            }
        }
    }
}