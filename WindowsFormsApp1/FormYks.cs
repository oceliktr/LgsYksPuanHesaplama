using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using WindowsFormsApp1.Library;
using WindowsFormsApp1.Model;
using Application = System.Windows.Forms.Application;
using DataTable = System.Data.DataTable;
// ReSharper disable ArrangeRedundantParentheses

namespace WindowsFormsApp1
{
    public partial class FormYks : Form
    {
        private int saat;
        private int dakika;
        private int saniye = 1;

        private string sinavAdi = "2020  YILI SON SINIF ÖĞRENCİLERİ TYT-AYT-YKS PUANLARI";

        private readonly List<YksSonuc> sonucList = new List<YksSonuc>();
        private string seciliDizin;
        public FormYks()
        {
            InitializeComponent();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("http://osmancelik.com.tr");
        }

        private void FormYks_Load(object sender, EventArgs e)
        {//background nesnesi sağlıklı çalışması için gerekli. 
            CheckForIllegalCrossThreadCalls = false;
        }

        private void btnGozat_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.ShowNewFolderButton = true; //yeni klasör oluşturmayı aç
                folderDialog.RootFolder = Environment.SpecialFolder.Desktop;
                folderDialog.SelectedPath = Environment.SpecialFolder.Desktop
                    .ToString(); //başlangıç dizini programın bulunduğu dizin => AppDomain.CurrentDomain.BaseDirectory
                folderDialog.Description = @"Karnelerin saklanacağı dizini seçiniz.";
                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    seciliDizin = folderDialog.SelectedPath;
                    btnBasla.Enabled = true;
                }
                folderDialog.Dispose();
            }

            label2.Text = "Hedef Dizin: " + seciliDizin;
        }

        private void btnBasla_Click(object sender, EventArgs e)
        {

            timer1.Enabled = true; //sayaç başlasın

            backgroundWorker1.RunWorkerAsync();

        }
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            btnBasla.Enabled = false;
            btnGozat.Enabled = false;

            OgrenciSonuclari();

            bool islemDevam = MukerrerKayitKontrol();
            if (islemDevam == false) //mükerre kayıt varsa dur
            {
                label4.Text = "Bitti.";
                GecenSureyiDurdur();
                btnBasla.Enabled = true;
                btnGozat.Enabled = true;
                return;
            }
            List<YksSonuc> ilIlceOrtalamalari = IlveIlceleriHesapla();
            List<YksSonuc> okulOrtalamalari = OkullariHesapla();

            ExcelTablosunaAktar(sonucList, okulOrtalamalari, ilIlceOrtalamalari);

            progressBar1.Value = 0;
            label4.Text = "Bitti.";

            btnBasla.Enabled = true;
            btnGozat.Enabled = true;
            GecenSureyiDurdur();
        }

        private void OgrenciSonuclari()
        {
            DizinIslemleri dizinIslemleri = new DizinIslemleri();
            List<DosyaInfo> dosyalar = dizinIslemleri.DizindekiDosyalariListele(seciliDizin);
            progressBar1.Maximum = dosyalar.Count;
            progressBar1.Value = 0;

            foreach (DosyaInfo dosya in dosyalar)
            {
                string filePath = dosya.DizinAdresi + @"\" + dosya.DosyaAdi;
                
                progressBar1.Value += 1;
                label4.Text = "Dosya sayısı : " + progressBar1.Value + "/" + dosyalar.Count + " ...veriler toplanıyor.";

                DataTable table = ExcelUtil.ExcelToDataTable(filePath);

                if (table.TableName == PuanTurleri.OkulAytTestSonucListesi.ToString())
                {
                    string[] okulAdiDizi = table.Rows[1][5].ToString().Replace("YKS DOĞRU YANLIŞ SAYILARI\n", "").Replace(" 2020  YILI SON SINIF ÖĞRENCİLERİ TEST SONUÇLARI LİSTESİ", "").Split('(');
                    string okulAdi = okulAdiDizi[0];
                    string ilceAdi = okulAdiDizi[1].Replace("ERZURUM - ", "").Replace(")", "");

                    for (int satir = 1; satir <= table.Rows.Count; satir += 1)
                    {
                        try
                        {
                            string tcKimlik = table.Rows[satir + 10][1].ToString();
                            string ogrenciAdi = table.Rows[satir + 10][4].ToString();
                           
                            int turkDiliDogru = table.Rows[satir + 10][6].ToInt32();
                            int turkDiliYanlis = table.Rows[satir + 10][7].ToInt32();
                            decimal turkDiliNet = table.Rows[satir + 10][9].ToDecimal();
                            bool turkDiliGirdi = (turkDiliDogru+ turkDiliYanlis) != 0;

                            int tarih1Dogru = table.Rows[satir + 10][10].ToInt32();
                            int tarih1Yanlis = table.Rows[satir + 10][11].ToInt32();
                            decimal tarih1Net = table.Rows[satir + 10][13].ToDecimal();
                            bool tarih1Girdi = (tarih1Dogru + tarih1Yanlis) != 0;

                            int cografya1Dogru = table.Rows[satir + 10][14].ToInt32();
                            int cografya1Yanlis = table.Rows[satir + 10][15].ToInt32();
                            decimal cografya1Net = table.Rows[satir + 10][17].ToDecimal();
                            bool cografya1Girdi = (cografya1Dogru + cografya1Yanlis) != 0;

                            int tarih2Dogru = table.Rows[satir + 10][18].ToInt32();
                            int tarih2Yanlis = table.Rows[satir + 10][19].ToInt32();
                            decimal tarih2Net = table.Rows[satir + 10][21].ToDecimal();
                            bool tarih2Girdi = (tarih2Dogru + tarih2Yanlis) != 0;

                            int cografya2Dogru = table.Rows[satir + 10][22].ToInt32();
                            int cografya2Yanlis = table.Rows[satir + 10][23].ToInt32();
                            decimal cografya2Net = table.Rows[satir + 10][25].ToDecimal();
                            bool cografya2Girdi = (cografya2Dogru + cografya2Yanlis) != 0;

                            int felsefeDogru = table.Rows[satir + 10][26].ToInt32();
                            int felsefeYanlis = table.Rows[satir + 10][27].ToInt32();
                            decimal felsefeNet = table.Rows[satir + 10][29].ToDecimal();
                            bool felsefeGirdi = (felsefeDogru + felsefeYanlis) != 0;

                            int dinDogru = table.Rows[satir + 10][30].ToInt32();
                            int dinYanlis = table.Rows[satir + 10][31].ToInt32();
                            decimal dinNet = table.Rows[satir + 10][33].ToDecimal();
                            bool dinGirdi = (dinDogru + dinYanlis) != 0;

                            int matAytDogru = table.Rows[satir + 10][36].ToInt32();
                            int matAytYanlis = table.Rows[satir + 10][37].ToInt32();
                            decimal matAytNet = table.Rows[satir + 10][39].ToDecimal();
                            bool mataytGirdi = (matAytDogru + matAytYanlis) != 0;

                            int fizDogru = table.Rows[satir + 10][41].ToInt32();
                            int fizYanlis = table.Rows[satir + 10][42].ToInt32();
                            decimal fizNet = table.Rows[satir + 10][44].ToDecimal();
                            bool fizGirdi = (fizDogru + fizYanlis) != 0;

                            int kimyaDogru = table.Rows[satir + 10][45].ToInt32();
                            int kimyaYanlis = table.Rows[satir + 10][46].ToInt32();
                            decimal kimyaNet = table.Rows[satir + 10][48].ToDecimal();
                            bool kimyaGirdi = (kimyaDogru + kimyaYanlis) != 0;

                            int biyolojiDogru = table.Rows[satir + 10][49].ToInt32();
                            int biyolojiYanlis = table.Rows[satir + 10][50].ToInt32();
                            decimal biyolojiNet = table.Rows[satir + 10][52].ToDecimal();
                            bool biyolojiGirdi = (biyolojiDogru + biyolojiYanlis) != 0;

                            decimal toplamAytNet = biyolojiNet + kimyaNet + fizNet + matAytNet + dinNet + felsefeNet +
                                                cografya2Net + tarih2Net + cografya1Net + tarih1Net + turkDiliNet;

                            YksSonuc sonuc = new YksSonuc
                            {
                                Kategori = table.TableName,
                                OkulAdi = okulAdi,
                                IlceAdi = ilceAdi,
                                Tckimlik = tcKimlik,
                                AdiSoyadi = ogrenciAdi,
                                TurkDiliDogru = turkDiliDogru,
                                TurkDiliYanlis = turkDiliYanlis,
                                TurkDiliNet = turkDiliNet,
                                TurkDiliGirdi = turkDiliGirdi,
                                Tarih1Dogru = tarih1Dogru,
                                Tarih1Yanlis = tarih1Yanlis,
                                Tarih1Net = tarih1Net,
                                Tarih1Girdi = tarih1Girdi,
                                Cog1Dogru = cografya1Dogru,
                                Cog1Yanlis = cografya1Yanlis,
                                Cog1Net = cografya1Net,
                                Cog1Girdi = cografya1Girdi,
                                Tarih2Dogru = tarih2Dogru,
                                Tarih2Yanlis = tarih2Yanlis,
                                Tarih2Net = tarih2Net,
                                Tarih2Girdi = tarih2Girdi,
                                Cog2Dogru = cografya2Dogru,
                                Cog2Yanlis = cografya2Yanlis,
                                Cog2Net = cografya2Net,
                                Cog2Girdi = cografya2Girdi,
                                FelsefeDogru = felsefeDogru,
                                FelsefeYanlis = felsefeYanlis,
                                FelsefeNet = felsefeNet,
                                FelsefeGirdi = felsefeGirdi,
                                DinDogru = dinDogru,
                                DinYanlis = dinYanlis,
                                DinNet = dinNet,
                                DinGirdi = dinGirdi,
                                MatDogruAyt = matAytDogru,
                                MatYanlisAyt = matAytYanlis,
                                MatNetAyt = matAytNet,
                                MatAytGirdi = mataytGirdi,
                                FizikDogru = fizDogru,
                                FizikYanlis = fizYanlis,
                                FizikNet = fizNet,
                                FizikGirdi = fizGirdi,
                                KimyaDogru = kimyaDogru,
                                KimyaYanlis = kimyaYanlis,
                                KimyaNet = kimyaNet,
                                KimyaGirdi = kimyaGirdi,
                                BiyolojiDogru = biyolojiDogru,
                                BiyolojiYanlis = biyolojiYanlis,
                                BiyolojiNet = biyolojiNet,
                                BiyolojiGirdi = biyolojiGirdi,
                                ToplamAytNet = toplamAytNet
                            };

                            sonucList.Add(sonuc);

                        }
                        catch (Exception)
                        {
                            //
                        }
                        Application.DoEvents();
                    }
                }
                if (table.TableName == PuanTurleri.OkulYksTestSonucListesi.ToString())
                {
                    string[] okulAdiDizi = table.Rows[1][5].ToString()
                        .Replace("YKS DOĞRU YANLIŞ SAYILARI\n", "")
                        .Replace(" 2020  YILI SON SINIF ÖĞRENCİLERİ TEST SONUÇLARI LİSTESİ", "").Split('(');
                    string okulAdi = okulAdiDizi[0];
                    string ilceAdi = okulAdiDizi[1].Replace("ERZURUM - ", "").Replace(")", "");

                    for (int satir = 1; satir <= table.Rows.Count; satir += 1)
                    {
                        try
                        {
                            string tcKimlik = table.Rows[satir + 10][1].ToString();
                            string ogrenciAdi = table.Rows[satir + 10][4].ToString();

                            int turkceDogru = table.Rows[satir + 10][6].ToInt32();
                            int turkceYanlis = table.Rows[satir + 10][7].ToInt32();
                            decimal turkceNet = table.Rows[satir + 10][9].ToDecimal();
                            bool turkceGirdi = (turkceDogru + turkceYanlis) != 0;

                            int sosyalDogru = table.Rows[satir + 10][10].ToInt32();
                            int sosyalYanlis = table.Rows[satir + 10][11].ToInt32();
                            decimal sosyalNet = table.Rows[satir + 10][13].ToDecimal();
                            bool sosyalGirdi = (sosyalDogru + sosyalYanlis) != 0;

                            int matTytDogru = table.Rows[satir + 10][14].ToInt32();
                            int matTytYanlis = table.Rows[satir + 10][15].ToInt32();
                            decimal matTytNet = table.Rows[satir + 10][17].ToDecimal();
                            bool matTytGirdi = (matTytDogru + matTytYanlis) != 0;

                            int fenDogru = table.Rows[satir + 10][18].ToInt32();
                            int fenYanlis = table.Rows[satir + 10][19].ToInt32();
                            decimal fenNet = table.Rows[satir + 10][21].ToDecimal();
                            bool fenGirdi = (fenDogru + fenYanlis) != 0;

                            decimal toplamTytNet = fenNet + matTytNet + sosyalNet + turkceNet;

                            YksSonuc sonuc = new YksSonuc
                            {
                                Kategori = table.TableName,
                                OkulAdi = okulAdi,
                                IlceAdi = ilceAdi,
                                Tckimlik = tcKimlik,
                                AdiSoyadi = ogrenciAdi,
                                TurkceDogru = turkceDogru,
                                TurkceYanlis = turkceYanlis,
                                TurkceNet = turkceNet,
                                TurkceGirdi = turkceGirdi,
                                SosyalBDogru = sosyalDogru,
                                SosyalBYanlis = sosyalYanlis,
                                SosyalBNet = sosyalNet,
                                SosyalBGirdi = sosyalGirdi,
                                MatDogruTyt = matTytDogru,
                                MatYanlisTyt = matTytYanlis,
                                MatNetTyt = matTytNet,
                                MatTytGirdi = matTytGirdi,
                                FenDogru = fenDogru,
                                FenYanlis = fenYanlis,
                                FenNet = fenNet,
                                FenGirdi = fenGirdi,
                                ToplamTytNet = toplamTytNet
                            };

                            sonucList.Add(sonuc);

                        }
                        catch (Exception)
                        {
                            //
                        }
                        Application.DoEvents();
                    }
                }

                if (table.TableName == PuanTurleri.OkulYksPuanlariListesi.ToString())
                {
                    string[] okulAdiDizi = table.Rows[1][5].ToString()
                        .Replace("YKS YERLEŞTİRME PUANLARI ve BAŞARI SIRLARI\n", "")
                        .Replace("YKS PUANLAR ve BAŞARI SIRLARI\n", "")
                        .Replace(" 2020  YILI SON SINIF ÖĞRENCİLERİ PUAN LİSTESİ", "").Split('(');
                    string okulAdi = okulAdiDizi[0];
                    string ilceAdi = okulAdiDizi[1].Replace("ERZURUM - ", "").Replace(")", "");

                    for (int satir = 1; satir <= table.Rows.Count; satir += 1)
                    {
                        try
                        {
                            string tcKimlik = table.Rows[satir + 9][1].ToString();
                            string ogrenciAdi = table.Rows[satir + 9][3].ToString();

                            decimal tytPuan = table.Rows[satir + 9][6].ToDecimal();
                            decimal sayisalPuan = table.Rows[satir + 9][8].ToDecimal();
                            decimal sozelPuan = table.Rows[satir + 9][10].ToDecimal();
                            decimal esitAgirlikPuan = table.Rows[satir + 9][12].ToDecimal();
                            decimal yabanciDilPuan = table.Rows[satir + 9][14].ToDecimal();

                            YksSonuc sonuc = new YksSonuc
                            {
                                Kategori = table.TableName,
                                OkulAdi = okulAdi,
                                IlceAdi = ilceAdi,
                                Tckimlik = tcKimlik,
                                AdiSoyadi = ogrenciAdi,
                                TYTPuanYuzde = tytPuan,
                                SayisalPuanYuzde = sayisalPuan,
                                SozelPuanYuzde = sozelPuan,
                                EsitAgirlikPuanYuzde = esitAgirlikPuan,
                                YabanciDilPuanYuzde = yabanciDilPuan
                            };

                            sonucList.Add(sonuc);
                            
                        }
                        catch (Exception)
                        {
                            //
                        }
                        Application.DoEvents();
                    }
                }
                if (table.TableName == PuanTurleri.OkulYksYerlestirmePuanlariListe.ToString())
                {
                    string[] okulAdiDizi = table.Rows[1][5].ToString()
                        .Replace("YKS YERLEŞTİRME PUANLARI ve BAŞARI SIRLARI\n", "")
                        .Replace("YKS PUANLAR ve BAŞARI SIRLARI\n", "")
                        .Replace(" 2020  YILI SON SINIF ÖĞRENCİLERİ PUAN LİSTESİ", "").Split('(');
                    string okulAdi = okulAdiDizi[0];
                    string ilceAdi = okulAdiDizi[1].Replace("ERZURUM - ", "").Replace(")", "");

                    for (int satir = 1; satir <= table.Rows.Count; satir += 1)
                    {
                        try
                        {
                            string tcKimlik = table.Rows[satir + 9][1].ToString();
                            string ogrenciAdi = table.Rows[satir + 9][3].ToString();

                            decimal tytPuan = table.Rows[satir + 9][6].ToDecimal();
                            decimal sayisalPuan = table.Rows[satir + 9][8].ToDecimal();
                            decimal sozelPuan = table.Rows[satir + 9][10].ToDecimal();
                            decimal esitAgirlikPuan = table.Rows[satir + 9][12].ToDecimal();
                            decimal yabanciDilPuan = table.Rows[satir + 9][14].ToDecimal();

                            YksSonuc sonuc = new YksSonuc
                            {
                                Kategori = table.TableName,
                                OkulAdi = okulAdi,
                                IlceAdi = ilceAdi,
                                Tckimlik = tcKimlik,
                                AdiSoyadi = ogrenciAdi,
                                TYTPuanYerl = tytPuan,
                                SayisalPuanYerl = sayisalPuan,
                                SozelPuanYerl = sozelPuan,
                                EsitAgirlikPuanYerl = esitAgirlikPuan,
                                YabanciDilPuanYerl = yabanciDilPuan
                            };

                            sonucList.Add(sonuc);

                        }
                        catch (Exception)
                        {
                            //
                        }
                        Application.DoEvents();
                    }
                }
               
                Application.DoEvents();
            }
            //dosyaları okuma bitti
        }

        private bool MukerrerKayitKontrol()
        {
            int islemSayisi = sonucList.Count;
            progressBar1.Maximum = islemSayisi;
            int a = 0;
            progressBar1.Value = 0;

            label4.Text = "Mükerrer kayıt kontrol ediliyor.";
            //Bilgileri geçici bellekte tutar.
            List<YksSonuc> mukerrerGeciciBellek = new List<YksSonuc>();
            List<YksSonuc> mukerrerList = new List<YksSonuc>();

            foreach (var item in sonucList)
            {
                a++;
                progressBar1.Value = a;

                var kontrol = mukerrerGeciciBellek.Find(x => x.Tckimlik == item.Tckimlik && x.Kategori == item.Kategori);
                if (kontrol == null)
                {
                    mukerrerGeciciBellek.Add(item); //bellekte yok ise ekle.
                }
                else
                {
                    MessageBox.Show("Mükkerrer Okullar\n" + item.Kategori + " - " + item.IlceAdi + " - " + item.OkulAdi + " - " + item.Tckimlik + " - " + item.AdiSoyadi + "\n");
                    mukerrerList.Add(item); //mükerrer ise ekle.
                }
            }

            if (mukerrerList.Count > 0)
            {
                string s = "";
                var mukerrerOkullar = mukerrerList.GroupBy(x => x.OkulAdi).Select(x => x.First()).ToList();
                foreach (var m in mukerrerOkullar)
                {
                    s += m.Kategori + " - " + m.IlceAdi + " - " + m.OkulAdi + " - " + m.Tckimlik + " - " + m.AdiSoyadi + "\n";

                }

                MessageBox.Show("Mükkerrer Okullar\n" + s);

                return false;
            }
            progressBar1.Value = 0;
            return true;
        }

        private List<YksSonuc> IlveIlceleriHesapla()
        {
            List<YksSonuc> ilIlceOrtalamalariList = new List<YksSonuc>();

            List<YksSonuc> ilceler = sonucList.GroupBy(x => x.IlceAdi).Select(x => x.First()).ToList();
            int islemSayisi = ilceler.Count;
            progressBar1.Maximum = islemSayisi;
            progressBar1.Value = 0;
            int a = 0;

            //ilçe ortalamaları
            foreach (var ilce in ilceler)
            {
                var ilceVerisi = sonucList.Where(x => x.IlceAdi == ilce.IlceAdi);


                int ogrenciSayisi = ilceVerisi.GroupBy(x => x.Tckimlik).Count();

                a++;
                progressBar1.Value = a;

                label4.Text = "İlçe ortalamaları hesaplanıyor." + a + "/" + islemSayisi;

                var ilceAytVerisi = ilceVerisi.Where(x => x.Kategori == PuanTurleri.OkulAytTestSonucListesi.ToString());
                var ilceTytVerisi = ilceVerisi.Where(x => x.Kategori == PuanTurleri.OkulYksTestSonucListesi.ToString());
                var ilceYerlestirmeVerisi = ilceVerisi.Where(x => x.Kategori == PuanTurleri.OkulYksYerlestirmePuanlariListe.ToString());
                var ilceYuzdelikVerisi = ilceVerisi.Where(x => x.Kategori == PuanTurleri.OkulYksPuanlariListesi.ToString());

                decimal toplamAytNet = ilceAytVerisi.Where(x => x.ToplamAytNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilceAytVerisi.Where(x => x.ToplamAytNet != 0).Sum(x => x.ToplamAytNet) / ilceAytVerisi.Where(x => x.ToplamAytNet != 0).GroupBy(x => x.Tckimlik).Count();
                decimal turkDiliNetAyt = ilceAytVerisi.Where(x => x.TurkDiliNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilceAytVerisi.Where(x => x.TurkDiliNet != 0).Sum(x => x.TurkDiliNet) / ilceAytVerisi.Where(x => x.TurkDiliNet != 0).GroupBy(x => x.Tckimlik).Count();
                decimal tarih1NetAyt = ilceAytVerisi.Where(x => x.Tarih1Net != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilceAytVerisi.Where(x => x.Tarih1Net != 0).Sum(x => x.Tarih1Net) / ilceAytVerisi.Where(x => x.Tarih1Net != 0).GroupBy(x => x.Tckimlik).Count();
                decimal cogr1NetAyt = ilceAytVerisi.Where(x => x.Cog1Net != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilceAytVerisi.Where(x => x.Cog1Net != 0).Sum(x => x.Cog1Net) / ilceAytVerisi.Where(x => x.Cog1Net != 0).GroupBy(x => x.Tckimlik).Count();
                decimal tarih2NetAyt = ilceAytVerisi.Where(x => x.Tarih2Net != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilceAytVerisi.Where(x => x.Tarih2Net != 0).Sum(x => x.Tarih2Net) / ilceAytVerisi.Where(x => x.Tarih2Net != 0).GroupBy(x => x.Tckimlik).Count();
                decimal cogr2NetAyt = ilceAytVerisi.Where(x => x.Cog2Net != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilceAytVerisi.Where(x => x.Cog2Net != 0).Sum(x => x.Cog2Net) / ilceAytVerisi.Where(x => x.Cog2Net != 0).GroupBy(x => x.Tckimlik).Count();
                decimal felsefeNetAyt = ilceAytVerisi.Where(x => x.FelsefeNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilceAytVerisi.Where(x => x.FelsefeNet != 0).Sum(x => x.FelsefeNet) / ilceAytVerisi.Where(x => x.FelsefeNet != 0).GroupBy(x => x.Tckimlik).Count();
                decimal dinNetAyt = ilceAytVerisi.Where(x => x.DinNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilceAytVerisi.Where(x => x.DinNet != 0).Sum(x => x.DinNet) / ilceAytVerisi.Where(x => x.DinNet != 0).GroupBy(x => x.Tckimlik).Count();
                decimal matNetAyt = ilceAytVerisi.Where(x => x.MatNetAyt != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilceAytVerisi.Where(x => x.MatNetAyt != 0).Sum(x => x.MatNetAyt) / ilceAytVerisi.Where(x => x.MatNetAyt != 0).GroupBy(x => x.Tckimlik).Count();
                decimal fizikNetAyt = ilceAytVerisi.Where(x => x.FizikNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilceAytVerisi.Where(x => x.FizikNet != 0).Sum(x => x.FizikNet) / ilceAytVerisi.Where(x => x.FizikNet != 0).GroupBy(x => x.Tckimlik).Count();
                decimal kimyaNetAyt = ilceAytVerisi.Where(x => x.KimyaNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilceAytVerisi.Where(x => x.KimyaNet != 0).Sum(x => x.KimyaNet) / ilceAytVerisi.Where(x => x.KimyaNet != 0).GroupBy(x => x.Tckimlik).Count();
                decimal biyolojiNetAyt = ilceAytVerisi.Where(x => x.BiyolojiNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilceAytVerisi.Where(x => x.BiyolojiNet != 0).Sum(x => x.BiyolojiNet) / ilceAytVerisi.Where(x => x.BiyolojiNet != 0).GroupBy(x => x.Tckimlik).Count();

                decimal toplamTytNet = ilceTytVerisi.Where(x => x.ToplamTytNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilceTytVerisi.Where(x => x.ToplamTytNet != 0).Sum(x => x.ToplamTytNet) / ilceTytVerisi.Where(x => x.ToplamTytNet != 0).GroupBy(x => x.Tckimlik).Count();
                decimal turkceNetTyt = ilceTytVerisi.Where(x => x.TurkceNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilceTytVerisi.Where(x => x.TurkceNet != 0).Sum(x => x.TurkceNet) / ilceTytVerisi.Where(x => x.TurkceNet != 0).GroupBy(x => x.Tckimlik).Count();
                decimal sosyalNetTyt = ilceTytVerisi.Where(x => x.SosyalBNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilceTytVerisi.Where(x => x.SosyalBNet != 0).Sum(x => x.SosyalBNet) / ilceTytVerisi.Where(x => x.SosyalBNet != 0).GroupBy(x => x.Tckimlik).Count();
                decimal matNetTyt = ilceTytVerisi.Where(x => x.MatNetTyt != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilceTytVerisi.Where(x => x.MatNetTyt != 0).Sum(x => x.MatNetTyt) / ilceTytVerisi.Where(x => x.MatNetTyt != 0).GroupBy(x => x.Tckimlik).Count();
                decimal fenNetTyt = ilceTytVerisi.Where(x => x.FenNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilceTytVerisi.Where(x => x.FenNet != 0).Sum(x => x.FenNet) / ilceTytVerisi.Where(x => x.FenNet != 0).GroupBy(x => x.Tckimlik).Count();

                decimal toplamTytYuzdelikPuani = ilceYuzdelikVerisi.Where(x => x.TYTPuanYuzde != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilceYuzdelikVerisi.Where(x => x.TYTPuanYuzde != 0).Sum(x => x.TYTPuanYuzde) / ilceYuzdelikVerisi.Where(x => x.TYTPuanYuzde != 0).GroupBy(x => x.Tckimlik).Count();
                decimal toplamSayisalYuzdelikPuani = ilceYuzdelikVerisi.Where(x => x.SayisalPuanYuzde != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilceYuzdelikVerisi.Where(x => x.SayisalPuanYuzde != 0).Sum(x => x.SayisalPuanYuzde) / ilceYuzdelikVerisi.Where(x => x.SayisalPuanYuzde != 0).GroupBy(x => x.Tckimlik).Count();
                decimal toplamSozelYuzdelikPuani = ilceYuzdelikVerisi.Where(x => x.SozelPuanYuzde != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilceYuzdelikVerisi.Where(x => x.SozelPuanYuzde != 0).Sum(x => x.SozelPuanYuzde) / ilceYuzdelikVerisi.Where(x => x.SozelPuanYuzde != 0).GroupBy(x => x.Tckimlik).Count();
                decimal toplamEsitAYuzdelikPuani = ilceYuzdelikVerisi.Where(x => x.EsitAgirlikPuanYuzde != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilceYuzdelikVerisi.Where(x => x.EsitAgirlikPuanYuzde != 0).Sum(x => x.EsitAgirlikPuanYuzde) / ilceYuzdelikVerisi.Where(x => x.EsitAgirlikPuanYuzde != 0).GroupBy(x => x.Tckimlik).Count();
                decimal toplamYDilYuzdelikPuani = ilceYuzdelikVerisi.Where(x => x.YabanciDilPuanYuzde != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilceYuzdelikVerisi.Where(x => x.YabanciDilPuanYuzde != 0).Sum(x => x.YabanciDilPuanYuzde) / ilceYuzdelikVerisi.Where(x => x.YabanciDilPuanYuzde != 0).GroupBy(x => x.Tckimlik).Count();


                decimal toplamTytYerlestirmePuani = ilceYerlestirmeVerisi.Where(x => x.TYTPuanYerl != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilceYerlestirmeVerisi.Where(x => x.TYTPuanYerl != 0).Sum(x => x.TYTPuanYerl) / ilceYerlestirmeVerisi.Where(x => x.TYTPuanYerl != 0).GroupBy(x => x.Tckimlik).Count();
                decimal toplamSayisalYerlestirmePuani = ilceYerlestirmeVerisi.Where(x => x.SayisalPuanYerl != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilceYerlestirmeVerisi.Where(x => x.SayisalPuanYerl != 0).Sum(x => x.SayisalPuanYerl) / ilceYerlestirmeVerisi.Where(x => x.SayisalPuanYerl != 0).GroupBy(x => x.Tckimlik).Count();
                decimal toplamSozelYerlestirmePuani = ilceYerlestirmeVerisi.Where(x => x.SozelPuanYerl != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilceYerlestirmeVerisi.Where(x => x.SozelPuanYerl != 0).Sum(x => x.SozelPuanYerl) / ilceYerlestirmeVerisi.Where(x => x.SozelPuanYerl != 0).GroupBy(x => x.Tckimlik).Count();
                decimal toplamEsitAYerlestirmePuani = ilceYerlestirmeVerisi.Where(x => x.EsitAgirlikPuanYerl != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilceYerlestirmeVerisi.Where(x => x.EsitAgirlikPuanYerl != 0).Sum(x => x.EsitAgirlikPuanYerl) / ilceYerlestirmeVerisi.Where(x => x.EsitAgirlikPuanYerl != 0).GroupBy(x => x.Tckimlik).Count();
                decimal toplamYDilYerlestirmePuani = ilceYerlestirmeVerisi.Where(x => x.YabanciDilPuanYerl != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilceYerlestirmeVerisi.Where(x => x.YabanciDilPuanYerl != 0).Sum(x => x.YabanciDilPuanYerl) / ilceYerlestirmeVerisi.Where(x => x.YabanciDilPuanYerl != 0).GroupBy(x => x.Tckimlik).Count();


                YksSonuc sonuc = new YksSonuc
                {
                    IlceAdi = ilce.IlceAdi,
                    OgrenciSayisi = ogrenciSayisi,
                    ToplamAytNet = toplamAytNet,
                    ToplamTytNet = toplamTytNet,

                    TurkDiliNet = turkDiliNetAyt,
                    Tarih1Net = tarih1NetAyt,
                    Cog1Net = cogr1NetAyt,
                    Tarih2Net = tarih2NetAyt,
                    Cog2Net = cogr2NetAyt,
                    FelsefeNet = felsefeNetAyt,
                    DinNet = dinNetAyt,
                    MatNetAyt = matNetAyt,
                    FizikNet = fizikNetAyt,
                    KimyaNet = kimyaNetAyt,
                    BiyolojiNet = biyolojiNetAyt,

                    TurkceNet = turkceNetTyt,
                    SosyalBNet = sosyalNetTyt,
                    MatNetTyt = matNetTyt,
                    FenNet = fenNetTyt,

                    TYTPuanYuzde = toplamTytYuzdelikPuani,
                    SozelPuanYuzde = toplamSozelYuzdelikPuani,
                    SayisalPuanYuzde = toplamSayisalYuzdelikPuani,
                    EsitAgirlikPuanYuzde = toplamEsitAYuzdelikPuani,
                    YabanciDilPuanYuzde = toplamYDilYuzdelikPuani,

                    TYTPuanYerl = toplamTytYerlestirmePuani,
                    SozelPuanYerl = toplamSozelYerlestirmePuani,
                    SayisalPuanYerl = toplamSayisalYerlestirmePuani,
                    EsitAgirlikPuanYerl = toplamEsitAYerlestirmePuani,
                    YabanciDilPuanYerl = toplamYDilYerlestirmePuani
                };

                ilIlceOrtalamalariList.Add(sonuc);
            }

            //il ortalaması

            var ilAytVerisi = sonucList.Where(x => x.Kategori == PuanTurleri.OkulAytTestSonucListesi.ToString());
            var ilTytVerisi = sonucList.Where(x => x.Kategori == PuanTurleri.OkulYksTestSonucListesi.ToString());
            var ilVerisiYuzdelik = sonucList.Where(x => x.Kategori == PuanTurleri.OkulYksPuanlariListesi.ToString());
            var ilVerisiYerlestirme = sonucList.Where(x => x.Kategori == PuanTurleri.OkulYksYerlestirmePuanlariListe.ToString());

            int ilOgrenciSayisi = sonucList.GroupBy(x => x.Tckimlik).Count();

            decimal toplamIlAytNet = sonucList.Where(x => x.ToplamAytNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilAytVerisi.Where(x => x.ToplamAytNet != 0).Sum(x => x.ToplamAytNet) / sonucList.Where(x => x.ToplamAytNet != 0).GroupBy(x => x.Tckimlik).Count();
            decimal turkDiliIlNetAyt = sonucList.Where(x => x.TurkDiliNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilAytVerisi.Where(x => x.TurkDiliNet != 0).Sum(x => x.TurkDiliNet) / sonucList.Where(x => x.TurkDiliNet != 0).GroupBy(x => x.Tckimlik).Count();
            decimal tarih1IlNetAyt = sonucList.Where(x => x.Tarih1Net != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilAytVerisi.Where(x => x.Tarih1Net != 0).Sum(x => x.Tarih1Net) / sonucList.Where(x => x.Tarih1Net != 0).GroupBy(x => x.Tckimlik).Count();
            decimal cogr1IlNetAyt = sonucList.Where(x => x.Cog1Net != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilAytVerisi.Where(x => x.Cog1Net != 0).Sum(x => x.Cog1Net) / sonucList.Where(x => x.Cog1Net != 0).GroupBy(x => x.Tckimlik).Count();
            decimal tarih2IlNetAyt = sonucList.Where(x => x.Tarih2Net != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilAytVerisi.Where(x => x.Tarih2Net != 0).Sum(x => x.Tarih2Net) / sonucList.Where(x => x.Tarih2Net != 0).GroupBy(x => x.Tckimlik).Count();
            decimal cogr2IlNetAyt = sonucList.Where(x => x.Cog2Net != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilAytVerisi.Where(x => x.Cog2Net != 0).Sum(x => x.Cog2Net) / sonucList.Where(x => x.Cog2Net != 0).GroupBy(x => x.Tckimlik).Count();
            decimal felsefeIlNetAyt = sonucList.Where(x => x.FelsefeNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilAytVerisi.Where(x => x.FelsefeNet != 0).Sum(x => x.FelsefeNet) / sonucList.Where(x => x.FelsefeNet != 0).GroupBy(x => x.Tckimlik).Count();
            decimal dinIlNetAyt = sonucList.Where(x => x.DinNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilAytVerisi.Where(x => x.DinNet != 0).Sum(x => x.DinNet) / sonucList.Where(x => x.DinNet != 0).GroupBy(x => x.Tckimlik).Count();
            decimal matIlNetAyt = sonucList.Where(x => x.MatNetAyt != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilAytVerisi.Where(x => x.MatNetAyt != 0).Sum(x => x.MatNetAyt) / sonucList.Where(x => x.MatNetAyt != 0).GroupBy(x => x.Tckimlik).Count();
            decimal fizikIlNetAyt = sonucList.Where(x => x.FizikNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilAytVerisi.Where(x => x.FizikNet != 0).Sum(x => x.FizikNet) / sonucList.Where(x => x.FizikNet != 0).GroupBy(x => x.Tckimlik).Count();
            decimal kimyaIlNetAyt = sonucList.Where(x => x.KimyaNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilAytVerisi.Where(x => x.KimyaNet != 0).Sum(x => x.KimyaNet) / sonucList.Where(x => x.KimyaNet != 0).GroupBy(x => x.Tckimlik).Count();
            decimal biyolojiIlNetAyt = sonucList.Where(x => x.BiyolojiNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilAytVerisi.Where(x => x.BiyolojiNet != 0).Sum(x => x.BiyolojiNet) / sonucList.Where(x => x.BiyolojiNet != 0).GroupBy(x => x.Tckimlik).Count();

            decimal toplamIlTytNet = sonucList.Where(x => x.ToplamTytNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilTytVerisi.Where(x => x.ToplamTytNet != 0).Sum(x => x.ToplamTytNet) / sonucList.Where(x => x.ToplamTytNet != 0).GroupBy(x => x.Tckimlik).Count();
            decimal turkceIlNetTyt = sonucList.Where(x => x.TurkceNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilTytVerisi.Where(x => x.TurkceNet != 0).Sum(x => x.TurkceNet) / sonucList.Where(x => x.TurkceNet != 0).GroupBy(x => x.Tckimlik).Count();
            decimal sosyalIlNetTyt = sonucList.Where(x => x.SosyalBNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilTytVerisi.Where(x => x.SosyalBNet != 0).Sum(x => x.SosyalBNet) / sonucList.Where(x => x.SosyalBNet != 0).GroupBy(x => x.Tckimlik).Count();
            decimal matIlNetTyt = sonucList.Where(x => x.MatNetTyt != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilTytVerisi.Where(x => x.MatNetTyt != 0).Sum(x => x.MatNetTyt) / sonucList.Where(x => x.MatNetTyt != 0).GroupBy(x => x.Tckimlik).Count();
            decimal fenIlNetTyt = sonucList.Where(x => x.FenNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilTytVerisi.Where(x => x.FenNet != 0).Sum(x => x.FenNet) / sonucList.Where(x => x.FenNet != 0).GroupBy(x => x.Tckimlik).Count();

            decimal toplamTytIlYuzdelikPuani = sonucList.Where(x => x.TYTPuanYuzde != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilVerisiYuzdelik.Where(x => x.TYTPuanYuzde != 0).Sum(x => x.TYTPuanYuzde) / sonucList.Where(x => x.TYTPuanYuzde != 0).GroupBy(x => x.Tckimlik).Count();
            decimal toplamSayisalIlYuzdelikPuani = sonucList.Where(x => x.SayisalPuanYuzde != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilVerisiYuzdelik.Where(x => x.SayisalPuanYuzde != 0).Sum(x => x.SayisalPuanYuzde) / sonucList.Where(x => x.SayisalPuanYuzde != 0).GroupBy(x => x.Tckimlik).Count();
            decimal toplamSozelIlYuzdelikPuani = sonucList.Where(x => x.SozelPuanYuzde != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilVerisiYuzdelik.Where(x => x.SozelPuanYuzde != 0).Sum(x => x.SozelPuanYuzde) / sonucList.Where(x => x.SozelPuanYuzde != 0).GroupBy(x => x.Tckimlik).Count();
            decimal toplamEsitAIlYuzdelikPuani = sonucList.Where(x => x.EsitAgirlikPuanYuzde != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilVerisiYuzdelik.Where(x => x.EsitAgirlikPuanYuzde != 0).Sum(x => x.EsitAgirlikPuanYuzde) / sonucList.Where(x => x.EsitAgirlikPuanYuzde != 0).GroupBy(x => x.Tckimlik).Count();
            decimal toplamYDilIlYuzdelikPuani = sonucList.Where(x => x.YabanciDilPuanYuzde != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilVerisiYuzdelik.Where(x => x.YabanciDilPuanYuzde != 0).Sum(x => x.YabanciDilPuanYuzde) / sonucList.Where(x => x.YabanciDilPuanYuzde != 0).GroupBy(x => x.Tckimlik).Count();

            decimal toplamTytIlYerlestirmePuani = sonucList.Where(x => x.TYTPuanYerl != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilVerisiYerlestirme.Where(x => x.TYTPuanYerl != 0).Sum(x => x.TYTPuanYerl) / sonucList.Where(x => x.TYTPuanYerl != 0).GroupBy(x => x.Tckimlik).Count();
            decimal toplamSayisalIlYerlestirmePuani = sonucList.Where(x => x.SayisalPuanYerl != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilVerisiYerlestirme.Where(x => x.SayisalPuanYerl != 0).Sum(x => x.SayisalPuanYerl) / sonucList.Where(x => x.SayisalPuanYerl != 0).GroupBy(x => x.Tckimlik).Count();
            decimal toplamSozelIlYerlestirmePuani = sonucList.Where(x => x.SozelPuanYerl != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilVerisiYerlestirme.Where(x => x.SozelPuanYerl != 0).Sum(x => x.SozelPuanYerl) / sonucList.Where(x => x.SozelPuanYerl != 0).GroupBy(x => x.Tckimlik).Count();
            decimal toplamEsitAIlYerlestirmePuani = sonucList.Where(x => x.EsitAgirlikPuanYerl != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilVerisiYerlestirme.Where(x => x.EsitAgirlikPuanYerl != 0).Sum(x => x.EsitAgirlikPuanYerl) / sonucList.Where(x => x.EsitAgirlikPuanYerl != 0).GroupBy(x => x.Tckimlik).Count();
            decimal toplamYDilIlYerlestirmePuani = sonucList.Where(x => x.YabanciDilPuanYerl != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : ilVerisiYerlestirme.Where(x => x.YabanciDilPuanYerl != 0).Sum(x => x.YabanciDilPuanYerl) / sonucList.Where(x => x.YabanciDilPuanYerl != 0).GroupBy(x => x.Tckimlik).Count();


            YksSonuc sonucIl = new YksSonuc
            {
                IlceAdi = "İl Ortalaması",
                OgrenciSayisi = ilOgrenciSayisi,
                ToplamAytNet = toplamIlAytNet,
                ToplamTytNet = toplamIlTytNet,

                TurkDiliNet = turkDiliIlNetAyt,
                Tarih1Net = tarih1IlNetAyt,
                Cog1Net = cogr1IlNetAyt,
                Tarih2Net = tarih2IlNetAyt,
                Cog2Net = cogr2IlNetAyt,
                FelsefeNet = felsefeIlNetAyt,
                DinNet = dinIlNetAyt,
                MatNetAyt = matIlNetAyt,
                FizikNet = fizikIlNetAyt,
                KimyaNet = kimyaIlNetAyt,
                BiyolojiNet = biyolojiIlNetAyt,

                TurkceNet = turkceIlNetTyt,
                SosyalBNet = sosyalIlNetTyt,
                MatNetTyt = matIlNetTyt,
                FenNet = fenIlNetTyt,

                TYTPuanYuzde = toplamTytIlYuzdelikPuani,
                SozelPuanYuzde = toplamSozelIlYuzdelikPuani,
                SayisalPuanYuzde = toplamSayisalIlYuzdelikPuani,
                EsitAgirlikPuanYuzde = toplamEsitAIlYuzdelikPuani,
                YabanciDilPuanYuzde = toplamYDilIlYuzdelikPuani,

                TYTPuanYerl = toplamTytIlYerlestirmePuani,
                SozelPuanYerl = toplamSozelIlYerlestirmePuani,
                SayisalPuanYerl = toplamSayisalIlYerlestirmePuani,
                EsitAgirlikPuanYerl = toplamEsitAIlYerlestirmePuani,
                YabanciDilPuanYerl = toplamYDilIlYerlestirmePuani,

            };

            ilIlceOrtalamalariList.Add(sonucIl);

            Application.DoEvents();
            progressBar1.Value = 0;
            return ilIlceOrtalamalariList;
        }
        private List<YksSonuc> OkullariHesapla()
        {
            List<YksSonuc> OkulOrtalamalariList = new List<YksSonuc>();

            List<YksSonuc> okullar = sonucList.GroupBy(x => new{ x.OkulAdi ,x.IlceAdi}).Select(x => x.First()).ToList();
            int islemSayisi = okullar.Count;
            progressBar1.Maximum = islemSayisi;
            progressBar1.Value = 0;
            int a = 0;

            //okul ortalamaları
            foreach (var okul in okullar)
            {
                var okulVerisi = sonucList.Where(x => x.OkulAdi == okul.OkulAdi);

                int ogrenciSayisi = okulVerisi.GroupBy(x => x.Tckimlik).Count();

                a++;
                progressBar1.Value = a;
                label4.Text = "Okul ortalamaları hesaplanıyor. " + a + "/" + islemSayisi;


                var okulAytVerisi = okulVerisi.Where(x => x.Kategori == PuanTurleri.OkulAytTestSonucListesi.ToString());
                var okulTytVerisi = okulVerisi.Where(x => x.Kategori == PuanTurleri.OkulYksTestSonucListesi.ToString());
                var okulVerisiYerlestirme = okulVerisi.Where(x => x.Kategori == PuanTurleri.OkulYksYerlestirmePuanlariListe.ToString());
                var okulVerisiYuzdelik = okulVerisi.Where(x => x.Kategori == PuanTurleri.OkulYksPuanlariListesi.ToString());

                decimal toplamAytNet = okulAytVerisi.Where(x => x.ToplamAytNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : okulAytVerisi.Where(x => x.ToplamAytNet != 0).Sum(x => x.ToplamAytNet) / okulAytVerisi.Where(x => x.ToplamAytNet != 0).GroupBy(x => x.Tckimlik).Count();
                decimal turkDiliNetAyt = okulAytVerisi.Where(x => x.TurkDiliNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : okulAytVerisi.Where(x => x.TurkDiliNet != 0).Sum(x => x.TurkDiliNet) / okulAytVerisi.Where(x => x.TurkDiliNet != 0).GroupBy(x => x.Tckimlik).Count();
                decimal tarih1NetAyt = okulAytVerisi.Where(x => x.Tarih1Net != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : okulAytVerisi.Where(x => x.Tarih1Net != 0).Sum(x => x.Tarih1Net) / okulAytVerisi.Where(x => x.Tarih1Net != 0).GroupBy(x => x.Tckimlik).Count();
                decimal cogr1NetAyt = okulAytVerisi.Where(x => x.Cog1Net != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : okulAytVerisi.Where(x => x.Cog1Net != 0).Sum(x => x.Cog1Net) / okulAytVerisi.Where(x => x.Cog1Net != 0).GroupBy(x => x.Tckimlik).Count();
                decimal tarih2NetAyt = okulAytVerisi.Where(x => x.Tarih2Net != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : okulAytVerisi.Where(x => x.Tarih2Net != 0).Sum(x => x.Tarih2Net) / okulAytVerisi.Where(x => x.Tarih2Net != 0).GroupBy(x => x.Tckimlik).Count();
                decimal cogr2NetAyt = okulAytVerisi.Where(x => x.Cog2Net != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : okulAytVerisi.Where(x => x.Cog2Net != 0).Sum(x => x.Cog2Net) / okulAytVerisi.Where(x => x.Cog2Net != 0).GroupBy(x => x.Tckimlik).Count();
                decimal felsefeNetAyt = okulAytVerisi.Where(x => x.FelsefeNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : okulAytVerisi.Where(x => x.FelsefeNet != 0).Sum(x => x.FelsefeNet) / okulAytVerisi.Where(x => x.FelsefeNet != 0).GroupBy(x => x.Tckimlik).Count();
                decimal dinNetAyt = okulAytVerisi.Where(x => x.DinNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : okulAytVerisi.Where(x => x.DinNet != 0).Sum(x => x.DinNet) / okulAytVerisi.Where(x => x.DinNet != 0).GroupBy(x => x.Tckimlik).Count();
                decimal matNetAyt = okulAytVerisi.Where(x => x.MatNetAyt != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : okulAytVerisi.Where(x => x.MatNetAyt != 0).Sum(x => x.MatNetAyt) / okulAytVerisi.Where(x => x.MatNetAyt != 0).GroupBy(x => x.Tckimlik).Count();
                decimal fizikNetAyt = okulAytVerisi.Where(x => x.FizikNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : okulAytVerisi.Where(x => x.FizikNet != 0).Sum(x => x.FizikNet) / okulAytVerisi.Where(x => x.FizikNet != 0).GroupBy(x => x.Tckimlik).Count();
                decimal kimyaNetAyt = okulAytVerisi.Where(x => x.KimyaNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : okulAytVerisi.Where(x => x.KimyaNet != 0).Sum(x => x.KimyaNet) / okulAytVerisi.Where(x => x.KimyaNet != 0).GroupBy(x => x.Tckimlik).Count();
                decimal biyolojiNetAyt = okulAytVerisi.Where(x => x.BiyolojiNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : okulAytVerisi.Where(x => x.BiyolojiNet != 0).Sum(x => x.BiyolojiNet) / okulAytVerisi.Where(x => x.BiyolojiNet != 0).GroupBy(x => x.Tckimlik).Count();

                decimal toplamTytNet = okulTytVerisi.Where(x => x.ToplamTytNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : okulTytVerisi.Where(x => x.ToplamTytNet != 0).Sum(x => x.ToplamTytNet) / okulTytVerisi.Where(x => x.ToplamTytNet != 0).GroupBy(x => x.Tckimlik).Count();
                decimal turkceNetTyt = okulTytVerisi.Where(x => x.TurkceNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : okulTytVerisi.Where(x => x.TurkceNet != 0).Sum(x => x.TurkceNet) / okulTytVerisi.Where(x => x.TurkceNet != 0).GroupBy(x => x.Tckimlik).Count();
                decimal sosyalNetTyt = okulTytVerisi.Where(x => x.SosyalBNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : okulTytVerisi.Where(x => x.SosyalBNet != 0).Sum(x => x.SosyalBNet) / okulTytVerisi.Where(x => x.SosyalBNet != 0).GroupBy(x => x.Tckimlik).Count();
                decimal matNetTyt = okulTytVerisi.Where(x => x.MatNetTyt != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : okulTytVerisi.Where(x => x.MatNetTyt != 0).Sum(x => x.MatNetTyt) / okulTytVerisi.Where(x => x.MatNetTyt != 0).GroupBy(x => x.Tckimlik).Count();
                decimal fenNetTyt = okulTytVerisi.Where(x => x.FenNet != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : okulTytVerisi.Where(x => x.FenNet != 0).Sum(x => x.FenNet) / okulTytVerisi.Where(x => x.FenNet != 0).GroupBy(x => x.Tckimlik).Count();

                decimal toplamTytYuzdelikPuani = okulVerisiYuzdelik.Where(x => x.TYTPuanYuzde != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : okulVerisiYuzdelik.Where(x => x.TYTPuanYuzde != 0).Sum(x => x.TYTPuanYuzde) / okulVerisiYuzdelik.Where(x => x.TYTPuanYuzde != 0).GroupBy(x => x.Tckimlik).Count();
                decimal toplamSayisalYuzdelikPuani = okulVerisiYuzdelik.Where(x => x.SayisalPuanYuzde != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : okulVerisiYuzdelik.Where(x => x.SayisalPuanYuzde != 0).Sum(x => x.SayisalPuanYuzde) / okulVerisiYuzdelik.Where(x => x.SayisalPuanYuzde != 0).GroupBy(x => x.Tckimlik).Count();
                decimal toplamSozelYuzdelikPuani = okulVerisiYuzdelik.Where(x => x.SozelPuanYuzde != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : okulVerisiYuzdelik.Where(x => x.SozelPuanYuzde != 0).Sum(x => x.SozelPuanYuzde) / okulVerisiYuzdelik.Where(x => x.SozelPuanYuzde != 0).GroupBy(x => x.Tckimlik).Count();
                decimal toplamEsitAYuzdelikPuani = okulVerisiYuzdelik.Where(x => x.EsitAgirlikPuanYuzde != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : okulVerisiYuzdelik.Where(x => x.EsitAgirlikPuanYuzde != 0).Sum(x => x.EsitAgirlikPuanYuzde) / okulVerisiYuzdelik.Where(x => x.EsitAgirlikPuanYuzde != 0).GroupBy(x => x.Tckimlik).Count();
                decimal toplamYDilYuzdelikPuani = okulVerisiYuzdelik.Where(x => x.YabanciDilPuanYuzde != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : okulVerisiYuzdelik.Where(x => x.YabanciDilPuanYuzde != 0).Sum(x => x.YabanciDilPuanYuzde) / okulVerisiYuzdelik.Where(x => x.YabanciDilPuanYuzde != 0).GroupBy(x => x.Tckimlik).Count();

                decimal toplamTytYerlestirmePuani = okulVerisiYerlestirme.Where(x => x.TYTPuanYerl != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : okulVerisiYerlestirme.Where(x => x.TYTPuanYerl != 0).Sum(x => x.TYTPuanYerl) / okulVerisiYerlestirme.Where(x => x.TYTPuanYerl != 0).GroupBy(x => x.Tckimlik).Count();
                decimal toplamSayisalYerlestirmePuani = okulVerisiYerlestirme.Where(x => x.SayisalPuanYerl != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : okulVerisiYerlestirme.Where(x => x.SayisalPuanYerl != 0).Sum(x => x.SayisalPuanYerl) / okulVerisiYerlestirme.Where(x => x.SayisalPuanYerl != 0).GroupBy(x => x.Tckimlik).Count();
                decimal toplamSozelYerlestirmePuani = okulVerisiYerlestirme.Where(x => x.SozelPuanYerl != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : okulVerisiYerlestirme.Where(x => x.SozelPuanYerl != 0).Sum(x => x.SozelPuanYerl) / okulVerisiYerlestirme.Where(x => x.SozelPuanYerl != 0).GroupBy(x => x.Tckimlik).Count();
                decimal toplamEsitAYerlestirmePuani = okulVerisiYerlestirme.Where(x => x.EsitAgirlikPuanYerl != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : okulVerisiYerlestirme.Where(x => x.EsitAgirlikPuanYerl != 0).Sum(x => x.EsitAgirlikPuanYerl) / okulVerisiYerlestirme.Where(x => x.EsitAgirlikPuanYerl != 0).GroupBy(x => x.Tckimlik).Count();
                decimal toplamYDilYerlestirmePuani = okulVerisiYerlestirme.Where(x => x.YabanciDilPuanYerl != 0).GroupBy(x => x.Tckimlik).Count() == 0 ? 0 : okulVerisiYerlestirme.Where(x => x.YabanciDilPuanYerl != 0).Sum(x => x.YabanciDilPuanYerl) / okulVerisiYerlestirme.Where(x => x.YabanciDilPuanYerl != 0).GroupBy(x => x.Tckimlik).Count();


                YksSonuc sonuc = new YksSonuc
                {
                    IlceAdi = okul.IlceAdi,
                    OkulAdi = okul.OkulAdi,
                    OgrenciSayisi = ogrenciSayisi,
                    ToplamAytNet = toplamAytNet,
                    ToplamTytNet = toplamTytNet,

                    TurkDiliNet = turkDiliNetAyt,
                    Tarih1Net = tarih1NetAyt,
                    Cog1Net = cogr1NetAyt,
                    Tarih2Net = tarih2NetAyt,
                    Cog2Net = cogr2NetAyt,
                    FelsefeNet = felsefeNetAyt,
                    DinNet = dinNetAyt,
                    MatNetAyt = matNetAyt,
                    FizikNet = fizikNetAyt,
                    KimyaNet = kimyaNetAyt,
                    BiyolojiNet = biyolojiNetAyt,

                    TurkceNet = turkceNetTyt,
                    SosyalBNet = sosyalNetTyt,
                    MatNetTyt = matNetTyt,
                    FenNet = fenNetTyt,

                    TYTPuanYuzde = toplamTytYuzdelikPuani,
                    SozelPuanYuzde = toplamSozelYuzdelikPuani,
                    SayisalPuanYuzde = toplamSayisalYuzdelikPuani,
                    EsitAgirlikPuanYuzde = toplamEsitAYuzdelikPuani,
                    YabanciDilPuanYuzde = toplamYDilYuzdelikPuani,

                    TYTPuanYerl = toplamTytYerlestirmePuani,
                    SozelPuanYerl = toplamSozelYerlestirmePuani,
                    SayisalPuanYerl = toplamSayisalYerlestirmePuani,
                    EsitAgirlikPuanYerl = toplamEsitAYerlestirmePuani,
                    YabanciDilPuanYerl = toplamYDilYerlestirmePuani
                };


                OkulOrtalamalariList.Add(sonuc);
            }
            Application.DoEvents();

            progressBar1.Value = 0;
            return OkulOrtalamalariList;
        }

        private void ExcelTablosunaAktar(List<YksSonuc> ogrenciXls, List<YksSonuc> okulXls, List<YksSonuc> ilIlceXls)
        {
            //excel baş

            string excelDosyaAdi = seciliDizin + "_Rapor_" + DateTime.Now.Ticks + ".xlsx";

            Microsoft.Office.Interop.Excel.Application aplicacion = new Microsoft.Office.Interop.Excel.Application();
            Workbook calismaKitabi = aplicacion.Workbooks.Add();


            ExcelOgrenciSayfasi(ogrenciXls, calismaKitabi);

            ExcelOkulSayfasi(okulXls, calismaKitabi);

            ExcelIlceSayfasi(ilIlceXls, calismaKitabi);


            calismaKitabi.SaveAs(excelDosyaAdi, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);

            calismaKitabi.Close(true);
            aplicacion.Quit();

            Process.Start(excelDosyaAdi);
        }
        private void ExcelIlceSayfasi(List<YksSonuc> ilceXls, Workbook calismaKitabi)
        {
            Sheets xlSheets = calismaKitabi.Sheets;
            var calismaSayfasi = (Worksheet)xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);

            calismaSayfasi.Name = "İLÇE";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 1, 1, 1, 30, 60); //başlık
            calismaSayfasi.Cells[1, 1] = sinavAdi;

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 4, 2, 15, 30);
            calismaSayfasi.Cells[2, 4] = "AYT";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 16, 3, 20, 30);
            calismaSayfasi.Cells[2, 16] = "TYT";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 21, 3, 25, 30);
            calismaSayfasi.Cells[2, 21] = "YÜZDELİK PUAN";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 26, 3, 30, 30);
            calismaSayfasi.Cells[2, 26] = "YERLEŞTİRME PUANI";



            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 1, 4, 1);
            calismaSayfasi.Cells[2, 1] = "No";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 2, 4, 2);
            calismaSayfasi.Cells[2, 2] = "İlçe Adı";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 3, 4, 3);
            calismaSayfasi.Cells[2, 3] = "Sınava Giren Öğr. S.";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 3, 4, 4, 4);
            calismaSayfasi.Cells[3, 4] = "Toplam Net Ort.";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 3, 5, 3, 7);
            calismaSayfasi.Cells[3, 5] = "Türk Dili Ve Edeb. - Sosyal Bil. -1";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 5, 4, 5);
            calismaSayfasi.Cells[4, 5] = "Türk Dili Ve Edeb. Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 6, 4, 6);
            calismaSayfasi.Cells[4, 6] = "Tarih -1 Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 7, 4, 7);
            calismaSayfasi.Cells[4, 7] = "Coğrafya -1 Net Ort.";


            ExcelUtil.HucreBirlesitir(calismaSayfasi, 3, 8, 3, 11);
            calismaSayfasi.Cells[3, 8] = "Sosyal Bilimler -2";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 8, 4, 8);
            calismaSayfasi.Cells[4, 8] = "Tarih -2 Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 9, 4, 9);
            calismaSayfasi.Cells[4, 9] = "Coğrafya -2 Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 10, 4, 10);
            calismaSayfasi.Cells[4, 10] = "Felsefe Grubu Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 11, 4, 11);
            calismaSayfasi.Cells[4, 11] = "Din Kültürü ve A.B. Net Ort.";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 3, 12, 3, 12);
            calismaSayfasi.Cells[3, 12] = "Matematik";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 12, 4, 12);
            calismaSayfasi.Cells[4, 12] = "Matematik Net Ort.";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 3, 13, 3, 15);
            calismaSayfasi.Cells[3, 13] = "Fen Bilimleri";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 13, 4, 13);
            calismaSayfasi.Cells[4, 13] = "Fizik Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 14, 4, 14);
            calismaSayfasi.Cells[4, 14] = "Kimya Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 15, 4, 15);
            calismaSayfasi.Cells[4, 15] = "Biyoloji Net Ort.";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 16, 4, 16);
            calismaSayfasi.Cells[4, 16] = "Toplam Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 17, 4, 17);
            calismaSayfasi.Cells[4, 17] = "Türkçe Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 18, 4, 18);
            calismaSayfasi.Cells[4, 18] = "Sosyal Bilimler Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 19, 4, 19);
            calismaSayfasi.Cells[4, 19] = "Temel Matematik Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 20, 4, 20);
            calismaSayfasi.Cells[4, 20] = "Fen Bilimleri Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 21, 4, 21);
            calismaSayfasi.Cells[4, 21] = "TYT Puan Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 22, 4, 22);
            calismaSayfasi.Cells[4, 22] = "Sayısal Puan Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 23, 4, 23);
            calismaSayfasi.Cells[4, 23] = "Sözel Puan Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 24, 4, 24);
            calismaSayfasi.Cells[4, 24] = "Eşit Ağırlık Puan Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 25, 4, 25);
            calismaSayfasi.Cells[4, 25] = "Yabancı Dil Puan Ort.";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 26, 4, 26);
            calismaSayfasi.Cells[4, 26] = "TYT Puan Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 27, 4, 27);
            calismaSayfasi.Cells[4, 27] = "Sayısal Puan Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 28, 4, 28);
            calismaSayfasi.Cells[4, 28] = "Sözel Puan Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 29, 4, 29);
            calismaSayfasi.Cells[4, 29] = "Eşit Ağırlık Puan Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 30, 4, 30);
            calismaSayfasi.Cells[4, 30] = "Yabancı Dil Puan Ort.";

            progressBar1.Maximum = ilceXls.Count;
            for (var i = 0; i < ilceXls.Count; i++)
            {
                label4.Text = $"İlçe sonuçları excele işleniyor {i + 1}/{ilceXls.Count}";
                progressBar1.Value = i;

                calismaSayfasi.Cells[5 + i, 1] = i + 1;
                calismaSayfasi.Cells[5 + i, 2] = ilceXls[i].IlceAdi;
                calismaSayfasi.Cells[5 + i, 3] = ilceXls[i].OgrenciSayisi;
                if (ilceXls[i].ToplamAytNet == 0) calismaSayfasi.Cells[5 + i, 4] = ""; else calismaSayfasi.Cells[5 + i, 4] = decimal.Round(ilceXls[i].ToplamAytNet, 3, MidpointRounding.AwayFromZero);
                if (ilceXls[i].TurkDiliNet == 0) calismaSayfasi.Cells[5 + i, 5] = ""; else calismaSayfasi.Cells[5 + i, 5] = decimal.Round(ilceXls[i].TurkDiliNet, 3, MidpointRounding.AwayFromZero);
                if (ilceXls[i].Tarih1Net == 0) calismaSayfasi.Cells[5 + i, 6] = ""; else calismaSayfasi.Cells[5 + i, 6] = decimal.Round(ilceXls[i].Tarih1Net, 3, MidpointRounding.AwayFromZero);
                if (ilceXls[i].Cog1Net == 0) calismaSayfasi.Cells[5 + i, 7] = ""; else calismaSayfasi.Cells[5 + i, 7] = decimal.Round(ilceXls[i].Cog1Net, 3, MidpointRounding.AwayFromZero);
                if (ilceXls[i].Tarih2Net == 0) calismaSayfasi.Cells[5 + i, 8] = ""; else calismaSayfasi.Cells[5 + i, 8] = decimal.Round(ilceXls[i].Tarih2Net, 3, MidpointRounding.AwayFromZero);
                if (ilceXls[i].Cog2Net == 0) calismaSayfasi.Cells[5 + i, 9] = ""; else calismaSayfasi.Cells[5 + i, 9] = decimal.Round(ilceXls[i].Cog2Net, 3, MidpointRounding.AwayFromZero);
                if (ilceXls[i].FelsefeNet == 0) calismaSayfasi.Cells[5 + i, 10] = ""; else calismaSayfasi.Cells[5 + i, 10] = decimal.Round(ilceXls[i].FelsefeNet, 3, MidpointRounding.AwayFromZero);
                if (ilceXls[i].DinNet == 0) calismaSayfasi.Cells[5 + i, 11] = ""; else calismaSayfasi.Cells[5 + i, 11] = decimal.Round(ilceXls[i].DinNet, 3, MidpointRounding.AwayFromZero);
                if (ilceXls[i].MatNetAyt == 0) calismaSayfasi.Cells[5 + i, 12] = ""; else calismaSayfasi.Cells[5 + i, 12] = decimal.Round(ilceXls[i].MatNetAyt, 3, MidpointRounding.AwayFromZero);
                if (ilceXls[i].FizikNet == 0) calismaSayfasi.Cells[5 + i, 13] = ""; else calismaSayfasi.Cells[5 + i, 13] = decimal.Round(ilceXls[i].FizikNet, 3, MidpointRounding.AwayFromZero);
                if (ilceXls[i].KimyaNet == 0) calismaSayfasi.Cells[5 + i, 14] = ""; else calismaSayfasi.Cells[5 + i, 14] = decimal.Round(ilceXls[i].KimyaNet, 3, MidpointRounding.AwayFromZero);
                if (ilceXls[i].BiyolojiNet == 0) calismaSayfasi.Cells[5 + i, 15] = ""; else calismaSayfasi.Cells[5 + i, 15] = decimal.Round(ilceXls[i].BiyolojiNet, 3, MidpointRounding.AwayFromZero);
                if (ilceXls[i].ToplamTytNet == 0) calismaSayfasi.Cells[5 + i, 16] = ""; else calismaSayfasi.Cells[5 + i, 16] = decimal.Round(ilceXls[i].ToplamTytNet, 3, MidpointRounding.AwayFromZero);
                if (ilceXls[i].TurkceNet == 0) calismaSayfasi.Cells[5 + i, 17] = ""; else calismaSayfasi.Cells[5 + i, 17] = decimal.Round(ilceXls[i].TurkceNet, 3, MidpointRounding.AwayFromZero);
                if (ilceXls[i].SosyalBNet == 0) calismaSayfasi.Cells[5 + i, 18] = ""; else calismaSayfasi.Cells[5 + i, 18] = decimal.Round(ilceXls[i].SosyalBNet, 3, MidpointRounding.AwayFromZero);
                if (ilceXls[i].MatNetTyt == 0) calismaSayfasi.Cells[5 + i, 19] = ""; else calismaSayfasi.Cells[5 + i, 19] = decimal.Round(ilceXls[i].MatNetTyt, 3, MidpointRounding.AwayFromZero);
                if (ilceXls[i].FenNet == 0) calismaSayfasi.Cells[5 + i, 20] = ""; else calismaSayfasi.Cells[5 + i, 20] = decimal.Round(ilceXls[i].FenNet, 3, MidpointRounding.AwayFromZero);
                if (ilceXls[i].TYTPuanYuzde == 0) calismaSayfasi.Cells[5 + i, 21] = ""; else calismaSayfasi.Cells[5 + i, 21] = decimal.Round(ilceXls[i].TYTPuanYuzde, 3, MidpointRounding.AwayFromZero);
                if (ilceXls[i].SayisalPuanYuzde == 0) calismaSayfasi.Cells[5 + i, 22] = ""; else calismaSayfasi.Cells[5 + i, 22] = decimal.Round(ilceXls[i].SayisalPuanYuzde, 3, MidpointRounding.AwayFromZero);
                if (ilceXls[i].SozelPuanYuzde == 0) calismaSayfasi.Cells[5 + i, 23] = ""; else calismaSayfasi.Cells[5 + i, 23] = decimal.Round(ilceXls[i].SozelPuanYuzde, 3, MidpointRounding.AwayFromZero);
                if (ilceXls[i].EsitAgirlikPuanYuzde == 0) calismaSayfasi.Cells[5 + i, 24] = ""; else calismaSayfasi.Cells[5 + i, 24] = decimal.Round(ilceXls[i].EsitAgirlikPuanYuzde, 3, MidpointRounding.AwayFromZero);
                if (ilceXls[i].YabanciDilPuanYuzde == 0) calismaSayfasi.Cells[5 + i, 25] = ""; else calismaSayfasi.Cells[5 + i, 25] = decimal.Round(ilceXls[i].YabanciDilPuanYuzde, 3, MidpointRounding.AwayFromZero);
                if (ilceXls[i].TYTPuanYerl == 0) calismaSayfasi.Cells[5 + i, 26] = ""; else calismaSayfasi.Cells[5 + i, 26] = decimal.Round(ilceXls[i].TYTPuanYerl, 3, MidpointRounding.AwayFromZero);
                if (ilceXls[i].SayisalPuanYerl == 0) calismaSayfasi.Cells[5 + i, 27] = ""; else calismaSayfasi.Cells[5 + i, 27] = decimal.Round(ilceXls[i].SayisalPuanYerl, 3, MidpointRounding.AwayFromZero);
                if (ilceXls[i].SozelPuanYerl == 0) calismaSayfasi.Cells[5 + i, 28] = ""; else calismaSayfasi.Cells[5 + i, 28] = decimal.Round(ilceXls[i].SozelPuanYerl, 3, MidpointRounding.AwayFromZero);
                if (ilceXls[i].EsitAgirlikPuanYerl == 0) calismaSayfasi.Cells[5 + i, 29] = ""; else calismaSayfasi.Cells[5 + i, 29] = decimal.Round(ilceXls[i].EsitAgirlikPuanYerl, 3, MidpointRounding.AwayFromZero);
                if (ilceXls[i].YabanciDilPuanYerl == 0) calismaSayfasi.Cells[5 + i, 30] = ""; else calismaSayfasi.Cells[5 + i, 30] = decimal.Round(ilceXls[i].YabanciDilPuanYerl, 3, MidpointRounding.AwayFromZero);
            }

            progressBar1.Value = 0;

            //başlık2 exceldeki ikinci satır net doğru yanlış bilgilerinin olduğu satır
            int satirGenisligi = 30;
            int satirBaslangici = 4;
            int kayitSayisi = ilceXls.Count;
            ExcelUtil.HucreSitili(calismaSayfasi, satirBaslangici, satirGenisligi, kayitSayisi);


        }
        private void ExcelOkulSayfasi(List<YksSonuc> okulXls, Workbook calismaKitabi)
        {
            Sheets xlSheets = calismaKitabi.Sheets;
            var calismaSayfasi = (Worksheet)xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);

            calismaSayfasi.Name = "OKUL";


            ExcelUtil.HucreBirlesitir(calismaSayfasi, 1, 1, 1, 31, 60); //başlık
            calismaSayfasi.Cells[1, 1] = sinavAdi;

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 5, 2, 16, 30);
            calismaSayfasi.Cells[2, 5] = "AYT";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 17, 3, 21, 30);
            calismaSayfasi.Cells[2, 17] = "TYT";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 22, 3, 26, 30);
            calismaSayfasi.Cells[2, 22] = "YÜZDELİK PUAN";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 27, 3, 31, 30);
            calismaSayfasi.Cells[2, 27] = "YERLEŞTİRME PUANI";



            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 1, 4, 1);
            calismaSayfasi.Cells[2, 1] = "No";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 2, 4, 2);
            calismaSayfasi.Cells[2, 2] = "İlçe Adı";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 3, 4, 3);
            calismaSayfasi.Cells[2, 3] = "Okul Adı";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 4, 4, 4);
            calismaSayfasi.Cells[2, 4] = "Sınava Giren Öğr. S.";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 3, 5, 4, 5);
            calismaSayfasi.Cells[3, 5] = "Toplam Net Ort.";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 3, 6, 3, 8);
            calismaSayfasi.Cells[3, 6] = "Türk Dili Ve Edeb. - Sosyal Bil. -1";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 6, 4, 6);
            calismaSayfasi.Cells[4, 6] = "Türk Dili Ve Edeb. Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 7, 4, 7);
            calismaSayfasi.Cells[4, 7] = "Tarih -1 Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 8, 4, 8);
            calismaSayfasi.Cells[4, 8] = "Coğrafya -1 Net Ort.";


            ExcelUtil.HucreBirlesitir(calismaSayfasi, 3, 9, 3, 12);
            calismaSayfasi.Cells[3, 9] = "Sosyal Bilimler -2";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 9, 4, 9);
            calismaSayfasi.Cells[4, 9] = "Tarih -2 Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 10, 4, 10);
            calismaSayfasi.Cells[4, 10] = "Coğrafya -2 Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 11, 4, 11);
            calismaSayfasi.Cells[4, 11] = "Felsefe Grubu Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 12, 4, 12);
            calismaSayfasi.Cells[4, 12] = "Din Kültürü ve A.B. Net Ort.";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 3, 13, 3, 13);
            calismaSayfasi.Cells[3, 13] = "Matematik";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 13, 4, 13);
            calismaSayfasi.Cells[4, 13] = "Matematik Net Ort.";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 3, 14, 3, 16);
            calismaSayfasi.Cells[3, 14] = "Fen Bilimleri";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 14, 4, 14);
            calismaSayfasi.Cells[4, 14] = "Fizik Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 15, 4, 15);
            calismaSayfasi.Cells[4, 15] = "Kimya Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 16, 4, 16);
            calismaSayfasi.Cells[4, 16] = "Biyoloji Net Ort.";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 17, 4, 17);
            calismaSayfasi.Cells[4, 17] = "Toplam Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 18, 4, 18);
            calismaSayfasi.Cells[4, 18] = "Türkçe Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 19, 4, 19);
            calismaSayfasi.Cells[4, 19] = "Sosyal Bilimler Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 20, 4, 20);
            calismaSayfasi.Cells[4, 20] = "Temel Matematik Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 21, 4, 21);
            calismaSayfasi.Cells[4, 21] = "Fen Bilimleri Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 22, 4, 22);
            calismaSayfasi.Cells[4, 22] = "TYT Puan Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 23, 4, 23);
            calismaSayfasi.Cells[4, 23] = "Sayısal Puan Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 24, 4, 24);
            calismaSayfasi.Cells[4, 24] = "Sözel Puan Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 25, 4, 25);
            calismaSayfasi.Cells[4, 25] = "Eşit Ağırlık Puan Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 26, 4, 26);
            calismaSayfasi.Cells[4, 26] = "Yabancı Dil Puan Ort.";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 27, 4, 27);
            calismaSayfasi.Cells[4, 27] = "TYT Puan Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 28, 4, 28);
            calismaSayfasi.Cells[4, 28] = "Sayısal Puan Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 29, 4, 29);
            calismaSayfasi.Cells[4, 29] = "Sözel Puan Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 30, 4, 30);
            calismaSayfasi.Cells[4, 30] = "Eşit Ağırlık Puan Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 31, 4, 31);
            calismaSayfasi.Cells[4, 31] = "Yabancı Dil Puan Ort.";

            progressBar1.Maximum = okulXls.Count;
            for (var i = 0; i < okulXls.Count; i++)
            {

                label4.Text = $"Okul sonuçları excele işleniyor. {i + 1}/{okulXls.Count}";
                progressBar1.Value = i;
                calismaSayfasi.Cells[5 + i, 1] = i + 1;
                calismaSayfasi.Cells[5 + i, 2] = okulXls[i].IlceAdi;
                calismaSayfasi.Cells[5 + i, 3] = okulXls[i].OkulAdi;
                calismaSayfasi.Cells[5 + i, 4] = okulXls[i].OgrenciSayisi;

                if (okulXls[i].ToplamAytNet == 0) calismaSayfasi.Cells[5 + i, 5] = ""; else calismaSayfasi.Cells[5 + i, 5] = decimal.Round(okulXls[i].ToplamAytNet, 3, MidpointRounding.AwayFromZero);
                if (okulXls[i].TurkDiliNet == 0) calismaSayfasi.Cells[5 + i, 6] = ""; else calismaSayfasi.Cells[5 + i, 6] = decimal.Round(okulXls[i].TurkDiliNet, 3, MidpointRounding.AwayFromZero);
                if (okulXls[i].Tarih1Net == 0) calismaSayfasi.Cells[5 + i, 7] = ""; else calismaSayfasi.Cells[5 + i, 7] = decimal.Round(okulXls[i].Tarih1Net, 3, MidpointRounding.AwayFromZero);
                if (okulXls[i].Cog1Net == 0) calismaSayfasi.Cells[5 + i, 8] = ""; else calismaSayfasi.Cells[5 + i, 8] = decimal.Round(okulXls[i].Cog1Net, 3, MidpointRounding.AwayFromZero);
                if (okulXls[i].Tarih2Net == 0) calismaSayfasi.Cells[5 + i, 9] = ""; else calismaSayfasi.Cells[5 + i, 9] = decimal.Round(okulXls[i].Tarih2Net, 3, MidpointRounding.AwayFromZero);
                if (okulXls[i].Cog2Net == 0) calismaSayfasi.Cells[5 + i, 10] = ""; else calismaSayfasi.Cells[5 + i, 10] = decimal.Round(okulXls[i].Cog2Net, 3, MidpointRounding.AwayFromZero);
                if (okulXls[i].FelsefeNet == 0) calismaSayfasi.Cells[5 + i, 11] = ""; else calismaSayfasi.Cells[5 + i, 11] = decimal.Round(okulXls[i].FelsefeNet, 3, MidpointRounding.AwayFromZero);
                if (okulXls[i].DinNet == 0) calismaSayfasi.Cells[5 + i, 12] = ""; else calismaSayfasi.Cells[5 + i, 12] = decimal.Round(okulXls[i].DinNet, 3, MidpointRounding.AwayFromZero);
                if (okulXls[i].MatNetAyt == 0) calismaSayfasi.Cells[5 + i, 13] = ""; else calismaSayfasi.Cells[5 + i, 13] = decimal.Round(okulXls[i].MatNetAyt, 3, MidpointRounding.AwayFromZero);
                if (okulXls[i].FizikNet == 0) calismaSayfasi.Cells[5 + i, 14] = ""; else calismaSayfasi.Cells[5 + i, 14] = decimal.Round(okulXls[i].FizikNet, 3, MidpointRounding.AwayFromZero);
                if (okulXls[i].KimyaNet == 0) calismaSayfasi.Cells[5 + i, 15] = ""; else calismaSayfasi.Cells[5 + i, 15] = decimal.Round(okulXls[i].KimyaNet, 3, MidpointRounding.AwayFromZero);
                if (okulXls[i].BiyolojiNet == 0) calismaSayfasi.Cells[5 + i, 16] = ""; else calismaSayfasi.Cells[5 + i, 16] = decimal.Round(okulXls[i].BiyolojiNet, 3, MidpointRounding.AwayFromZero);
                if (okulXls[i].ToplamTytNet == 0) calismaSayfasi.Cells[5 + i, 17] = ""; else calismaSayfasi.Cells[5 + i, 17] = decimal.Round(okulXls[i].ToplamTytNet, 3, MidpointRounding.AwayFromZero);
                if (okulXls[i].TurkceNet == 0) calismaSayfasi.Cells[5 + i, 18] = ""; else calismaSayfasi.Cells[5 + i, 18] = decimal.Round(okulXls[i].TurkceNet, 3, MidpointRounding.AwayFromZero);
                if (okulXls[i].SosyalBNet == 0) calismaSayfasi.Cells[5 + i, 19] = ""; else calismaSayfasi.Cells[5 + i, 19] = decimal.Round(okulXls[i].SosyalBNet, 3, MidpointRounding.AwayFromZero);
                if (okulXls[i].MatNetTyt == 0) calismaSayfasi.Cells[5 + i, 20] = ""; else calismaSayfasi.Cells[5 + i, 20] = decimal.Round(okulXls[i].MatNetTyt, 3, MidpointRounding.AwayFromZero);
                if (okulXls[i].FenNet == 0) calismaSayfasi.Cells[5 + i, 21] = ""; else calismaSayfasi.Cells[5 + i, 21] = decimal.Round(okulXls[i].FenNet, 3, MidpointRounding.AwayFromZero);
                if (okulXls[i].TYTPuanYuzde == 0) calismaSayfasi.Cells[5 + i, 22] = ""; else calismaSayfasi.Cells[5 + i, 22] = decimal.Round(okulXls[i].TYTPuanYuzde, 3, MidpointRounding.AwayFromZero);
                if (okulXls[i].SayisalPuanYuzde == 0) calismaSayfasi.Cells[5 + i, 23] = ""; else calismaSayfasi.Cells[5 + i, 23] = decimal.Round(okulXls[i].SayisalPuanYuzde, 3, MidpointRounding.AwayFromZero);
                if (okulXls[i].SozelPuanYuzde == 0) calismaSayfasi.Cells[5 + i, 24] = ""; else calismaSayfasi.Cells[5 + i, 24] = decimal.Round(okulXls[i].SozelPuanYuzde, 3, MidpointRounding.AwayFromZero);
                if (okulXls[i].EsitAgirlikPuanYuzde == 0) calismaSayfasi.Cells[5 + i, 25] = ""; else calismaSayfasi.Cells[5 + i, 25] = decimal.Round(okulXls[i].EsitAgirlikPuanYuzde, 3, MidpointRounding.AwayFromZero);
                if (okulXls[i].YabanciDilPuanYuzde == 0) calismaSayfasi.Cells[5 + i, 26] = ""; else calismaSayfasi.Cells[5 + i, 26] = decimal.Round(okulXls[i].YabanciDilPuanYuzde, 3, MidpointRounding.AwayFromZero);
                if (okulXls[i].TYTPuanYerl == 0) calismaSayfasi.Cells[5 + i, 27] = ""; else calismaSayfasi.Cells[5 + i, 27] = decimal.Round(okulXls[i].TYTPuanYerl, 3, MidpointRounding.AwayFromZero);
                if (okulXls[i].SayisalPuanYerl == 0) calismaSayfasi.Cells[5 + i, 28] = ""; else calismaSayfasi.Cells[5 + i, 28] = decimal.Round(okulXls[i].SayisalPuanYerl, 3, MidpointRounding.AwayFromZero);
                if (okulXls[i].SozelPuanYerl == 0) calismaSayfasi.Cells[5 + i, 29] = ""; else calismaSayfasi.Cells[5 + i, 29] = decimal.Round(okulXls[i].SozelPuanYerl, 3, MidpointRounding.AwayFromZero);
                if (okulXls[i].EsitAgirlikPuanYerl == 0) calismaSayfasi.Cells[5 + i, 30] = ""; else calismaSayfasi.Cells[5 + i, 30] = decimal.Round(okulXls[i].EsitAgirlikPuanYerl, 3, MidpointRounding.AwayFromZero);
                if (okulXls[i].YabanciDilPuanYerl == 0) calismaSayfasi.Cells[5 + i, 31] = ""; else calismaSayfasi.Cells[5 + i, 31] = decimal.Round(okulXls[i].YabanciDilPuanYerl, 3, MidpointRounding.AwayFromZero);

            }
            progressBar1.Value = 0;

            int satirGenisligi = 31;
            int satirBaslangici = 4;
            int kayitSayisi = okulXls.Count;
            ExcelUtil.HucreSitili(calismaSayfasi, satirBaslangici, satirGenisligi, kayitSayisi);
        }

        private void ExcelOgrenciSayfasi(List<YksSonuc> ogrenciXls, Workbook calismaKitabi)
        {
            //öğrenci listesini puana göre yeniden sıralama yap.

            Worksheet calismaSayfasi = (Worksheet)calismaKitabi.Worksheets.Item[1];

            calismaSayfasi.Name = "OGRENCI";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 1, 1, 1, 31, 60); //başlık
            calismaSayfasi.Cells[1, 1] = sinavAdi;

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 5, 2, 16, 30);
            calismaSayfasi.Cells[2, 5] = "AYT";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 17, 3, 21, 30);
            calismaSayfasi.Cells[2, 17] = "TYT";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 22, 3, 26, 30);
            calismaSayfasi.Cells[2, 22] = "YÜZDELİK PUAN";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 27, 3, 31, 30);
            calismaSayfasi.Cells[2, 27] = "YERLEŞTİRME PUANI";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 1, 4, 1);
            calismaSayfasi.Cells[2, 1] = "No";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 2, 4, 2);
            calismaSayfasi.Cells[2, 2] = "İlçe Adı";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 3, 4, 3);
            calismaSayfasi.Cells[2, 3] = "Okul Adı";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 4, 4, 4);
            calismaSayfasi.Cells[2, 4] = "Adı Soyadı";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 3, 5, 4, 5);
            calismaSayfasi.Cells[3, 5] = "Toplam Net Ort.";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 3, 6, 3, 8);
            calismaSayfasi.Cells[3, 6] = "Türk Dili Ve Edeb. - Sosyal Bil. -1";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 6, 4, 6);
            calismaSayfasi.Cells[4, 6] = "Türk Dili Ve Edeb. Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 7, 4, 7);
            calismaSayfasi.Cells[4, 7] = "Tarih -1 Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 8, 4, 8);
            calismaSayfasi.Cells[4, 8] = "Coğrafya -1 Net Ort.";


            ExcelUtil.HucreBirlesitir(calismaSayfasi, 3, 9, 3, 12);
            calismaSayfasi.Cells[3, 9] = "Sosyal Bilimler -2";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 9, 4, 9);
            calismaSayfasi.Cells[4, 9] = "Tarih -2 Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 10, 4, 10);
            calismaSayfasi.Cells[4, 10] = "Coğrafya -2 Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 11, 4, 11);
            calismaSayfasi.Cells[4, 11] = "Felsefe Grubu Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 12, 4, 12);
            calismaSayfasi.Cells[4, 12] = "Din Kültürü ve A.B. Net Ort.";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 3, 13, 3, 13);
            calismaSayfasi.Cells[3, 13] = "Matematik";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 13, 4, 13);
            calismaSayfasi.Cells[4, 13] = "Matematik Net Ort.";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 3, 14, 3, 16);
            calismaSayfasi.Cells[3, 14] = "Fen Bilimleri";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 14, 4, 14);
            calismaSayfasi.Cells[4, 14] = "Fizik Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 15, 4, 15);
            calismaSayfasi.Cells[4, 15] = "Kimya Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 16, 4, 16);
            calismaSayfasi.Cells[4, 16] = "Biyoloji Net Ort.";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 17, 4, 17);
            calismaSayfasi.Cells[4, 17] = "Toplam Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 18, 4, 18);
            calismaSayfasi.Cells[4, 18] = "Türkçe Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 19, 4, 19);
            calismaSayfasi.Cells[4, 19] = "Sosyal Bilimler Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 20, 4, 20);
            calismaSayfasi.Cells[4, 20] = "Temel Matematik Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 21, 4, 21);
            calismaSayfasi.Cells[4, 21] = "Fen Bilimleri Net Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 22, 4, 22);
            calismaSayfasi.Cells[4, 22] = "TYT Puan Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 23, 4, 23);
            calismaSayfasi.Cells[4, 23] = "Sayısal Puan Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 24, 4, 24);
            calismaSayfasi.Cells[4, 24] = "Sözel Puan Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 25, 4, 25);
            calismaSayfasi.Cells[4, 25] = "Eşit Ağırlık Puan Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 26, 4, 26);
            calismaSayfasi.Cells[4, 26] = "Yabancı Dil Puan Ort.";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 27, 4, 27);
            calismaSayfasi.Cells[4, 27] = "TYT Puan Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 28, 4, 28);
            calismaSayfasi.Cells[4, 28] = "Sayısal Puan Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 29, 4, 29);
            calismaSayfasi.Cells[4, 29] = "Sözel Puan Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 30, 4, 30);
            calismaSayfasi.Cells[4, 30] = "Eşit Ağırlık Puan Ort.";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 4, 31, 4, 31);
            calismaSayfasi.Cells[4, 31] = "Yabancı Dil Puan Ort.";

            int ogrenciSayisi = ogrenciXls.GroupBy(x => x.Tckimlik).Count();
            progressBar1.Maximum = ogrenciSayisi;

            List<YksSonuc> ogrenciList = sonucList.GroupBy(x => x.Tckimlik).Select(x => x.First()).ToList();

            int i = 0;
            foreach (var item in ogrenciList)
            {
                var ogrenci = sonucList.FirstOrDefault(x => x.Tckimlik == item.Tckimlik && x.Kategori == PuanTurleri.OkulYksTestSonucListesi.ToString());
                progressBar1.Value = i;
                label4.Text = $"Öğrenci sonuçları excele işleniyor {i + 1}/{ogrenciSayisi}";

                calismaSayfasi.Cells[5 + i, 1] = i + 1;
                calismaSayfasi.Cells[5 + i, 2] = ogrenci?.IlceAdi;
                calismaSayfasi.Cells[5 + i, 3] = ogrenci?.OkulAdi;
                calismaSayfasi.Cells[5 + i, 4] = ogrenci?.AdiSoyadi;

                ogrenci = sonucList.FirstOrDefault(x => x.Tckimlik == item.Tckimlik && x.Kategori == PuanTurleri.OkulAytTestSonucListesi.ToString());

                if (ogrenci != null)
                {
                    // ReSharper disable once SpecifyACultureInStringConversionExplicitly
                     calismaSayfasi.Cells[5 + i, 5] = decimal.Round(ogrenci.ToplamAytNet, 3, MidpointRounding.AwayFromZero);
                    if (ogrenci.TurkDiliNet == 0) calismaSayfasi.Cells[5 + i, 6] = ""; else calismaSayfasi.Cells[5 + i, 6] = decimal.Round(ogrenci.TurkDiliNet, 3, MidpointRounding.AwayFromZero);
                    if (ogrenci.Tarih1Net == 0) calismaSayfasi.Cells[5 + i, 7] = ""; else calismaSayfasi.Cells[5 + i, 7] = decimal.Round(ogrenci.Tarih1Net, 3, MidpointRounding.AwayFromZero);
                    if (ogrenci.Cog1Net == 0) calismaSayfasi.Cells[5 + i, 8] = ""; else calismaSayfasi.Cells[5 + i, 8] = decimal.Round(ogrenci.Cog1Net, 3, MidpointRounding.AwayFromZero);
                    if (ogrenci.Tarih2Net == 0) calismaSayfasi.Cells[5 + i, 9] = ""; else calismaSayfasi.Cells[5 + i, 9] = decimal.Round(ogrenci.Tarih2Net, 3, MidpointRounding.AwayFromZero);
                    if (ogrenci.Cog2Net == 0) calismaSayfasi.Cells[5 + i, 10] = ""; else calismaSayfasi.Cells[5 + i, 10] = decimal.Round(ogrenci.Cog2Net, 3, MidpointRounding.AwayFromZero);
                    if (ogrenci.FelsefeNet == 0) calismaSayfasi.Cells[5 + i, 11] = ""; else calismaSayfasi.Cells[5 + i, 11] = decimal.Round(ogrenci.FelsefeNet, 3, MidpointRounding.AwayFromZero);
                    if (ogrenci.DinNet == 0) calismaSayfasi.Cells[5 + i, 12] = ""; else calismaSayfasi.Cells[5 + i, 12] = decimal.Round(ogrenci.DinNet, 3, MidpointRounding.AwayFromZero);
                    if (ogrenci.MatNetAyt == 0) calismaSayfasi.Cells[5 + i, 13] = ""; else calismaSayfasi.Cells[5 + i, 13] = decimal.Round(ogrenci.MatNetAyt, 3, MidpointRounding.AwayFromZero);
                    if (ogrenci.FizikNet == 0) calismaSayfasi.Cells[5 + i, 14] = ""; else calismaSayfasi.Cells[5 + i, 14] = decimal.Round(ogrenci.FizikNet, 3, MidpointRounding.AwayFromZero);
                    if (ogrenci.KimyaNet == 0) calismaSayfasi.Cells[5 + i, 15] = ""; else calismaSayfasi.Cells[5 + i, 15] = decimal.Round(ogrenci.KimyaNet, 3, MidpointRounding.AwayFromZero);
                    if (ogrenci.BiyolojiNet == 0) calismaSayfasi.Cells[5 + i, 16] = ""; else calismaSayfasi.Cells[5 + i, 16] = decimal.Round(ogrenci.BiyolojiNet, 3, MidpointRounding.AwayFromZero);
                }

                ogrenci = sonucList.FirstOrDefault(x => x.Tckimlik == item.Tckimlik && x.Kategori == PuanTurleri.OkulYksTestSonucListesi.ToString());

                if (ogrenci != null)
                {
                    calismaSayfasi.Cells[5 + i, 17] = decimal.Round(ogrenci.ToplamTytNet, 3, MidpointRounding.AwayFromZero);
                    if (ogrenci.TurkceNet == 0) calismaSayfasi.Cells[5 + i, 18] = ""; else calismaSayfasi.Cells[5 + i, 18] = decimal.Round(ogrenci.TurkceNet, 3, MidpointRounding.AwayFromZero);
                    if (ogrenci.SosyalBNet == 0) calismaSayfasi.Cells[5 + i, 19] = ""; else calismaSayfasi.Cells[5 + i, 19] = decimal.Round(ogrenci.SosyalBNet, 3, MidpointRounding.AwayFromZero);
                    if (ogrenci.MatNetTyt == 0) calismaSayfasi.Cells[5 + i, 20] = ""; else calismaSayfasi.Cells[5 + i, 20] = decimal.Round(ogrenci.MatNetTyt, 3, MidpointRounding.AwayFromZero);
                    if (ogrenci.FenNet == 0) calismaSayfasi.Cells[5 + i, 21] = ""; else calismaSayfasi.Cells[5 + i, 21] = decimal.Round(ogrenci.FenNet, 3, MidpointRounding.AwayFromZero);
                }

                ogrenci = sonucList.FirstOrDefault(x => x.Tckimlik == item.Tckimlik && x.Kategori == PuanTurleri.OkulYksPuanlariListesi.ToString());

                if (ogrenci != null)
                {
                    if (ogrenci.TYTPuanYuzde == 0) calismaSayfasi.Cells[5 + i, 22] = ""; else calismaSayfasi.Cells[5 + i, 22] = decimal.Round(ogrenci.TYTPuanYuzde, 3, MidpointRounding.AwayFromZero);
                    if (ogrenci.SayisalPuanYuzde == 0) calismaSayfasi.Cells[5 + i, 23] = ""; else calismaSayfasi.Cells[5 + i, 23] = decimal.Round(ogrenci.SayisalPuanYuzde, 3, MidpointRounding.AwayFromZero);
                    if (ogrenci.SozelPuanYuzde == 0) calismaSayfasi.Cells[5 + i, 24] = ""; else calismaSayfasi.Cells[5 + i, 24] = decimal.Round(ogrenci.SozelPuanYuzde, 3, MidpointRounding.AwayFromZero);
                    if (ogrenci.EsitAgirlikPuanYuzde == 0) calismaSayfasi.Cells[5 + i, 25] = ""; else calismaSayfasi.Cells[5 + i, 25] = decimal.Round(ogrenci.EsitAgirlikPuanYuzde, 3, MidpointRounding.AwayFromZero);
                    if (ogrenci.YabanciDilPuanYuzde == 0) calismaSayfasi.Cells[5 + i, 26] = ""; else calismaSayfasi.Cells[5 + i, 26] = decimal.Round(ogrenci.YabanciDilPuanYuzde, 3, MidpointRounding.AwayFromZero);
                }

                ogrenci = sonucList.FirstOrDefault(x => x.Tckimlik == item.Tckimlik && x.Kategori == PuanTurleri.OkulYksYerlestirmePuanlariListe.ToString());

                if (ogrenci != null)
                {
                    if (ogrenci.TYTPuanYerl == 0) calismaSayfasi.Cells[5 + i, 27] = ""; else calismaSayfasi.Cells[5 + i, 27] = decimal.Round(ogrenci.TYTPuanYerl, 3, MidpointRounding.AwayFromZero);
                    if (ogrenci.SayisalPuanYerl == 0) calismaSayfasi.Cells[5 + i, 28] = ""; else calismaSayfasi.Cells[5 + i, 28] = decimal.Round(ogrenci.SayisalPuanYerl, 3, MidpointRounding.AwayFromZero);
                    if (ogrenci.SozelPuanYerl == 0) calismaSayfasi.Cells[5 + i, 29] = ""; else calismaSayfasi.Cells[5 + i, 29] = decimal.Round(ogrenci.SozelPuanYerl, 3, MidpointRounding.AwayFromZero);
                    if (ogrenci.EsitAgirlikPuanYerl == 0) calismaSayfasi.Cells[5 + i, 30] = ""; else calismaSayfasi.Cells[5 + i, 30] = decimal.Round(ogrenci.EsitAgirlikPuanYerl, 3, MidpointRounding.AwayFromZero);
                    if (ogrenci.YabanciDilPuanYerl == 0) calismaSayfasi.Cells[5 + i, 31] = ""; else calismaSayfasi.Cells[5 + i, 31] = decimal.Round(ogrenci.YabanciDilPuanYerl, 3, MidpointRounding.AwayFromZero);
                }

                i++;
            }
            progressBar1.Value = 0;

            //başlık2 exceldeki ikinci satır net doğru yanlış bilgilerinin olduğu satır
            int satirGenisligi = 31;
            int satirBaslangici = 4;
            ExcelUtil.HucreSitili(calismaSayfasi, satirBaslangici, satirGenisligi, ogrenciSayisi);


        }

        private void GecenSureyiDurdur()
        {
            timer1.Enabled = false; //geçen süreyi durdur
            saat = 0;
            dakika = 0;
            saniye = 0;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            lblGecenSure.Text = string.Format("Geçen süre: {0:D2}:{1:D2}:{2:D2}", saat, dakika, saniye);
            saniye++;
            if (saniye == 59)
            {
                saniye = 0;
                dakika++;
                if (dakika == 59)
                {
                    saat++;
                    dakika = 0;
                }
            }
        }

        private void FormYks_FormClosing(object sender, FormClosingEventArgs e)
        {
            FormGiris frm = new FormGiris();
            frm.Show();
        }
    }
}
