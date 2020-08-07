using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using WindowsFormsApp1.Library;
using WindowsFormsApp1.Model;
using Application = System.Windows.Forms.Application;
using DataTable = System.Data.DataTable;

namespace WindowsFormsApp1
{
    public partial class FormLgs : Form
    {
        private int saat;
        private int dakika;
        private int saniye;

        private string sinavAdi = "";
        private readonly List<LgsSonuc> sonucList = new List<LgsSonuc>();
        private string seciliDizin;
        public FormLgs()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            //background nesnesi sağlıklı çalışması için gerekli. 
            CheckForIllegalCrossThreadCalls = false;
        }

        private void btnBasla_Click(object sender, EventArgs e)
        {
            timer1.Enabled = true; //sayaç başlasın

            backgroundWorker1.RunWorkerAsync();

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

        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            btnBasla.Enabled = false;
            btnGozat.Enabled = false;

            //sonuçları dosyalardan oku ve sonucList e ekle
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
            var ilIlceOrtalamalari = IlveIlceleriHesapla();
            var okulOrtalamalari = OkullariHesapla();

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

                string ilceAdi = dosya.DizinAdresi.Replace(seciliDizin, "").Replace("\\", "");

                progressBar1.Value += 1;
                label4.Text = "Okul sayısı : " + progressBar1.Value + "/" + dosyalar.Count + " ...veriler toplanıyor.";

                DataTable table = ExcelUtil.ExcelToDataTable(filePath);


                for (int satir = 1; satir <= table.Rows.Count; satir += 18)
                {
                    string aciklama = table.Rows[satir + 5][11].ToString();
                    if (sinavAdi == "")
                        sinavAdi = table.Rows[17][0].ToString().ToUpper(CultureInfo.CurrentCulture);
                    //row 
                    string tcKimlik = table.Rows[satir - 1][7].ToString();
                    string adiSoyadi = table.Rows[satir][7].ToString();
                    string okulu = table.Rows[satir + 3][7].ToString().Replace("İ", "İ").Replace("ı", "ı");

                    int turkceDogru = table.Rows[satir + 5][6].ToInt32();
                    int turkceYanlis = table.Rows[satir + 5][8].ToInt32();
                    decimal turkceNet = turkceDogru - ((decimal)turkceYanlis / 3);

                    int matDogru = table.Rows[satir + 6][6].ToInt32();
                    int matYanlis = table.Rows[satir + 6][8].ToInt32();
                    decimal matNet = matDogru - ((decimal)matYanlis / 3);

                    int fenDogru = table.Rows[satir + 7][6].ToInt32();
                    int fenYanlis = table.Rows[satir + 7][8].ToInt32();
                    decimal fenNet = fenDogru - ((decimal)fenYanlis / 3);

                    int inkDogru = table.Rows[satir + 8][6].ToInt32();
                    int inkYanlis = table.Rows[satir + 8][8].ToInt32();
                    decimal inkNet = inkDogru - ((decimal)inkYanlis / 3);

                    int dinDogru = table.Rows[satir + 10][6].ToInt32();
                    int dinYanlis = table.Rows[satir + 10][8].ToInt32();
                    decimal dinNet = dinDogru - ((decimal)dinYanlis / 3);

                    int ingDogru = table.Rows[satir + 11][6].ToInt32();
                    int ingYanlis = table.Rows[satir + 11][8].ToInt32();
                    decimal ingNet = ingDogru - ((decimal)ingYanlis / 3);

                    decimal sinavPuani = table.Rows[satir + 14][6].ToDecimal();
                    decimal yuzdelik = table.Rows[satir + 12][3].ToString().Replace("Türkiye Geneli Yüzdelik Dilimi: %", "")
                        .ToDecimal();

                    //   okulu = Encoding.UTF8.GetString(Encoding.Default.GetBytes(okulu));
                    //  adiSoyadi = Encoding.UTF8.GetString(Encoding.Default.GetBytes(adiSoyadi));
                    //  aciklama = Encoding.UTF8.GetString(Encoding.Default.GetBytes(aciklama));
                    decimal toplamNet = (turkceNet + matNet + fenNet + inkNet + dinNet + ingNet);

                    LgsSonuc sonuc = new LgsSonuc
                    {
                        Tckimlik = tcKimlik.Trim(),
                        SinavPuani = sinavPuani,
                        YuzdelikDilim = yuzdelik,
                        ToplamNet = toplamNet,
                        TurkceDogru = turkceDogru,
                        TurkceYanlis = turkceYanlis,
                        TurkceNet = turkceNet,
                        MatDogru = matDogru,
                        MatYanlis = matYanlis,
                        MatNet = matNet,
                        FenDogru = fenDogru,
                        FenYanlis = fenYanlis,
                        FenNet = fenNet,
                        InkDogru = inkDogru,
                        InkYanlis = inkYanlis,
                        InkNet = inkNet,
                        DinDogru = dinDogru,
                        DinYanlis = dinYanlis,
                        DinNet = dinNet,
                        IngDogru = ingDogru,
                        IngYanlis = ingYanlis,
                        IngNet = ingNet,
                        Aciklama = aciklama,
                        AdiSoyadi = adiSoyadi,
                        IlceAdi = ilceAdi,
                        OkulAdi = okulu
                    };
                    sonucList.Add(sonuc);
                    //SonucManager manager = new SonucManager();
                    //manager.Insert(sonuc);

                    Application.DoEvents();
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
            List<LgsSonuc> mukerrerGeciciBellek = new List<LgsSonuc>();
            List<LgsSonuc> mukerrerList = new List<LgsSonuc>();

            foreach (var item in sonucList)
            {
                a++;
                progressBar1.Value = a;

                var kontrol = mukerrerGeciciBellek.Find(x => x.Tckimlik == item.Tckimlik);
                if (kontrol == null)
                {
                    mukerrerGeciciBellek.Add(item); //bellekte yok ise ekle.
                }
                else
                {
                    mukerrerList.Add(item); //bellekte yok ise ekle.
                }
            }

            if (mukerrerList.Count > 0)
            {
                string s = "";
                var mukerrerOkullar = mukerrerList.GroupBy(x => x.OkulAdi).Select(x => x.First()).ToList();
                foreach (var m in mukerrerOkullar)
                {
                    s += m.IlceAdi + "-" + m.OkulAdi + "\n";

                }

                MessageBox.Show("Mükkerrer Okullar\n" + s);

                return false;
            }
            progressBar1.Value = 0;
            return true;
        }
        private List<LgsSonuc> IlveIlceleriHesapla()
        {
            List<LgsSonuc> ilIlceOrtalamalariList = new List<LgsSonuc>();


            List<LgsSonuc> ilceler = sonucList.Where(x => x.Aciklama == "-").GroupBy(x => x.IlceAdi).Select(x => x.First()).ToList();
            int islemSayisi = ilceler.Count;
            progressBar1.Maximum = islemSayisi;
            progressBar1.Value = 0;
            int a = 0;

            //ilçe ortalamaları
            foreach (var ilce in ilceler)
            {
                var ilceVerisi = sonucList.Where(x => x.IlceAdi == ilce.IlceAdi && x.Aciklama == "-");


                int ogrenciSayisi = ilceVerisi.Count();

                a++;
                progressBar1.Value = a;

                label4.Text = "İlçe ortalamaları hesaplanıyor." + a + "/" + islemSayisi;

                int turkceDogru = ilceVerisi.Sum(x => x.TurkceDogru);
                int turkceYanlis = ilceVerisi.Sum(x => x.TurkceYanlis);
                decimal turkceNet = (turkceDogru - ((decimal)turkceYanlis / 3)) / ogrenciSayisi;

                int matDogru = ilceVerisi.Sum(x => x.MatDogru);
                int matYanlis = ilceVerisi.Sum(x => x.MatYanlis);
                decimal matNet = (matDogru - ((decimal)matYanlis / 3)) / ogrenciSayisi;

                int fenDogru = ilceVerisi.Sum(x => x.FenDogru);
                int fenYanlis = ilceVerisi.Sum(x => x.FenYanlis);
                decimal fenNet = (fenDogru - ((decimal)fenYanlis / 3)) / ogrenciSayisi;

                int inkDogru = ilceVerisi.Sum(x => x.InkDogru);
                int inkYanlis = ilceVerisi.Sum(x => x.IngYanlis);
                decimal inkNet = (inkDogru - ((decimal)inkYanlis / 3)) / ogrenciSayisi;

                int dinDogru = ilceVerisi.Sum(x => x.DinDogru);
                int dinYanlis = ilceVerisi.Sum(x => x.DinYanlis);
                decimal dinNet = (dinDogru - ((decimal)dinYanlis / 3)) / ogrenciSayisi;

                int ingDogru = ilceVerisi.Sum(x => x.IngDogru);
                int ingYanlis = ilceVerisi.Sum(x => x.IngYanlis);
                decimal ingNet = (ingDogru - ((decimal)ingYanlis / 3)) / ogrenciSayisi;


                decimal toplamNet = ilceVerisi.Sum(x => x.ToplamNet) / ogrenciSayisi;
                decimal sinavPuanOrtalamasi = ilceVerisi.Sum(x => x.SinavPuani) / ogrenciSayisi;

                LgsSonuc sonuc = new LgsSonuc
                {
                    ToplamNet = toplamNet,
                    SinavPuani = sinavPuanOrtalamasi,
                    TurkceDogru = turkceDogru,
                    TurkceYanlis = turkceYanlis,
                    TurkceNet = turkceNet,
                    MatDogru = matDogru,
                    MatYanlis = matYanlis,
                    MatNet = matNet,
                    FenDogru = fenDogru,
                    FenYanlis = fenYanlis,
                    FenNet = fenNet,
                    InkDogru = inkDogru,
                    InkYanlis = inkYanlis,
                    InkNet = inkNet,
                    DinDogru = dinDogru,
                    DinYanlis = dinYanlis,
                    DinNet = dinNet,
                    IngDogru = ingDogru,
                    IngYanlis = ingYanlis,
                    IngNet = ingNet,
                    IlceAdi = ilce.IlceAdi,
                    OgrenciSayisi = ogrenciSayisi
                };

                ilIlceOrtalamalariList.Add(sonuc);
            }

            //il ortalaması

            var ilVerisi = sonucList.Where(x => x.Aciklama == "-");

            int iLogrenciSayisi = ilVerisi.Count();

            int turkceIlDogru = ilVerisi.Sum(x => x.TurkceDogru);
            int turkceIlYanlis = ilVerisi.Sum(x => x.TurkceYanlis);
            decimal turkceIlNet = (turkceIlDogru - ((decimal)turkceIlYanlis / 3)) / iLogrenciSayisi;

            int matIlDogru = ilVerisi.Sum(x => x.MatDogru);
            int matIlYanlis = ilVerisi.Sum(x => x.MatYanlis);
            decimal matIlNet = (matIlDogru - ((decimal)matIlYanlis / 3)) / iLogrenciSayisi;

            int fenIlDogru = ilVerisi.Sum(x => x.FenDogru);
            int fenIlYanlis = ilVerisi.Sum(x => x.FenYanlis);
            decimal fenIlNet = (fenIlDogru - ((decimal)fenIlYanlis / 3)) / iLogrenciSayisi;

            int inkIlDogru = ilVerisi.Sum(x => x.InkDogru);
            int inkIlYanlis = ilVerisi.Sum(x => x.IngYanlis);
            decimal inkIlNet = (inkIlDogru - ((decimal)inkIlYanlis / 3)) / iLogrenciSayisi;

            int dinIlDogru = ilVerisi.Sum(x => x.DinDogru);
            int dinIlYanlis = ilVerisi.Sum(x => x.DinYanlis);
            decimal dinIlNet = (dinIlDogru - ((decimal)dinIlYanlis / 3)) / iLogrenciSayisi;

            int ingIlDogru = ilVerisi.Sum(x => x.IngDogru);
            int ingIlYanlis = ilVerisi.Sum(x => x.IngYanlis);
            decimal ingIlNet = (ingIlDogru - ((decimal)ingIlYanlis / 3)) / iLogrenciSayisi;


            decimal toplamIlNet = ilVerisi.Sum(x => x.ToplamNet) / iLogrenciSayisi;

            decimal sinavIlPuanOrtalamasi = ilVerisi.Sum(x => x.SinavPuani) / iLogrenciSayisi;

            LgsSonuc sonucIl = new LgsSonuc
            {
                ToplamNet = toplamIlNet,
                TurkceDogru = turkceIlDogru,
                TurkceYanlis = turkceIlYanlis,
                TurkceNet = turkceIlNet,
                MatDogru = matIlDogru,
                MatYanlis = matIlYanlis,
                MatNet = matIlNet,
                FenDogru = fenIlDogru,
                FenYanlis = fenIlYanlis,
                FenNet = fenIlNet,
                InkDogru = inkIlDogru,
                InkYanlis = inkIlYanlis,
                InkNet = inkIlNet,
                DinDogru = dinIlDogru,
                DinYanlis = dinIlYanlis,
                DinNet = dinIlNet,
                IngDogru = ingIlDogru,
                IngYanlis = ingIlYanlis,
                IngNet = ingIlNet,
                IlceAdi = "İl Ortalaması",
                OgrenciSayisi = iLogrenciSayisi,
                SinavPuani = sinavIlPuanOrtalamasi
            };

            ilIlceOrtalamalariList.Add(sonucIl);

            Application.DoEvents();
            progressBar1.Value = 0;
            return ilIlceOrtalamalariList;
        }

        private List<LgsSonuc> OkullariHesapla()
        {
            List<LgsSonuc> OkulOrtalamalariList = new List<LgsSonuc>();

            List<LgsSonuc> okullar = sonucList.Where(x => x.Aciklama == "-").GroupBy(x => new { x.IlceAdi, x.OkulAdi }).Select(x => x.First()).ToList();
            int islemSayisi = okullar.Count;
            progressBar1.Maximum = islemSayisi;
            progressBar1.Value = 0;
            int a = 0;

            //okul ortalamaları
            foreach (var okul in okullar)
            {
                var okulVerisi = sonucList.Where(x => x.OkulAdi == okul.OkulAdi && x.Aciklama == "-");

                int ogrenciSayisi = okulVerisi.Count();

                a++;
                progressBar1.Value = a;
                label4.Text = "Okul ortalamaları hesaplanıyor. " + a + "/" + islemSayisi;

                int turkceDogru = okulVerisi.Sum(x => x.TurkceDogru);
                int turkceYanlis = okulVerisi.Sum(x => x.TurkceYanlis);
                decimal turkceNet = (turkceDogru - ((decimal)turkceYanlis / 3)) / ogrenciSayisi;

                int matDogru = okulVerisi.Sum(x => x.MatDogru);
                int matYanlis = okulVerisi.Sum(x => x.MatYanlis);
                decimal matNet = (matDogru - ((decimal)matYanlis / 3)) / ogrenciSayisi;

                int fenDogru = okulVerisi.Sum(x => x.FenDogru);
                int fenYanlis = okulVerisi.Sum(x => x.FenYanlis);
                decimal fenNet = (fenDogru - ((decimal)fenYanlis / 3)) / ogrenciSayisi;

                int inkDogru = okulVerisi.Sum(x => x.InkDogru);
                int inkYanlis = okulVerisi.Sum(x => x.IngYanlis);
                decimal inkNet = (inkDogru - ((decimal)inkYanlis / 3)) / ogrenciSayisi;

                int dinDogru = okulVerisi.Sum(x => x.DinDogru);
                int dinYanlis = okulVerisi.Sum(x => x.DinYanlis);
                decimal dinNet = (dinDogru - ((decimal)dinYanlis / 3)) / ogrenciSayisi;

                int ingDogru = okulVerisi.Sum(x => x.IngDogru);
                int ingYanlis = okulVerisi.Sum(x => x.IngYanlis);
                decimal ingNet = (ingDogru - ((decimal)ingYanlis / 3)) / ogrenciSayisi;

                decimal toplamNet = okulVerisi.Sum(x => x.ToplamNet) / ogrenciSayisi;
                decimal sinavPuanOrtalamasi = okulVerisi.Sum(x => x.SinavPuani) / ogrenciSayisi;

                LgsSonuc sonuc = new LgsSonuc
                {
                    OgrenciSayisi = ogrenciSayisi,
                    SinavPuani = sinavPuanOrtalamasi,
                    ToplamNet = toplamNet,
                    TurkceDogru = turkceDogru,
                    TurkceYanlis = turkceYanlis,
                    TurkceNet = turkceNet,
                    MatDogru = matDogru,
                    MatYanlis = matYanlis,
                    MatNet = matNet,
                    FenDogru = fenDogru,
                    FenYanlis = fenYanlis,
                    FenNet = fenNet,
                    InkDogru = inkDogru,
                    InkYanlis = inkYanlis,
                    InkNet = inkNet,
                    DinDogru = dinDogru,
                    DinYanlis = dinYanlis,
                    DinNet = dinNet,
                    IngDogru = ingDogru,
                    IngYanlis = ingYanlis,
                    IngNet = ingNet,
                    IlceAdi = okul.IlceAdi,
                    OkulAdi = okul.OkulAdi
                };

                OkulOrtalamalariList.Add(sonuc);
            }
            Application.DoEvents();

            progressBar1.Value = 0;
            return OkulOrtalamalariList;
        }

        private void ExcelTablosunaAktar(List<LgsSonuc> ogrenciXls, List<LgsSonuc> okulXls, List<LgsSonuc> ilIlceXls)
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

            //calismaKitabi.Close(true);
            //aplicacion.Quit();

            // calismaKitabi.SaveAs(excelDosyaAdi, XlFileFormat.xlWorkbookNormal);
            calismaKitabi.Close(true);
            aplicacion.Quit();

            Process.Start(excelDosyaAdi);
        }
        private void ExcelOgrenciSayfasi(List<LgsSonuc> ogrenciXls, Workbook calismaKitabi)
        {
            //öğrenci listesini puana göre yeniden sıralama yap.
            ogrenciXls = ogrenciXls.Where(x => x.Aciklama == "-").OrderByDescending(x => x.SinavPuani).ToList();

            Worksheet calismaSayfasi = (Worksheet)calismaKitabi.Worksheets.Item[1];

            calismaSayfasi.Name = "OGRENCI";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 1, 1, 1, 25, 60);
            calismaSayfasi.Cells[1, 1] = sinavAdi;


            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 8, 2, 10);
            calismaSayfasi.Cells[2, 8] = "Türkçe";


            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 11, 2, 13);
            calismaSayfasi.Cells[2, 11] = "Matematik";


            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 14, 2, 16);
            calismaSayfasi.Cells[2, 14] = "Fen Bilimleri";


            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 17, 2, 19);
            calismaSayfasi.Cells[2, 17] = "İnkılap Tarihi";


            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 20, 2, 22);
            calismaSayfasi.Cells[2, 20] = "DKAB";


            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 23, 2, 25);
            calismaSayfasi.Cells[2, 23] = "Yabancı Dil";


            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 1, 3, 1);
            calismaSayfasi.Cells[2, 1] = "No";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 2, 3, 2);
            calismaSayfasi.Cells[2, 2] = "İlçe Adı";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 3, 3, 3);
            calismaSayfasi.Cells[2, 3] = "Kurum Adı";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 4, 3, 4);
            calismaSayfasi.Cells[2, 4] = "Adı Soyadı";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 5, 3, 5);
            calismaSayfasi.Cells[2, 5] = "Sınav Puanı";
            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 6, 3, 6);
            calismaSayfasi.Cells[2, 6] = "Yüzdelik Dilim";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 7, 3, 7);
            calismaSayfasi.Cells[2, 7] = "Toplam Net";

            calismaSayfasi.Cells[3, 8] = "D";//Türkçe
            calismaSayfasi.Cells[3, 9] = "Y";//Türkçe
            calismaSayfasi.Cells[3, 10] = "Net";//Türkçe
            calismaSayfasi.Cells[3, 11] = "D";//Matematik
            calismaSayfasi.Cells[3, 12] = "Y";//Matematik
            calismaSayfasi.Cells[3, 13] = "Net";//Matematik
            calismaSayfasi.Cells[3, 14] = "D";//Fen Bilimleri
            calismaSayfasi.Cells[3, 15] = "Y";//Fen Bilimleri
            calismaSayfasi.Cells[3, 16] = "Net";//Fen Bilimleri
            calismaSayfasi.Cells[3, 17] = "D";//İnkılap Tarihi
            calismaSayfasi.Cells[3, 18] = "Y";//İnkılap Tarihi
            calismaSayfasi.Cells[3, 19] = "Net";//İnkılap Tarihi
            calismaSayfasi.Cells[3, 20] = "D";  //DKAB
            calismaSayfasi.Cells[3, 21] = "Y";  //DKAB
            calismaSayfasi.Cells[3, 22] = "Net";//DKAB
            calismaSayfasi.Cells[3, 23] = "D";  //Yabancı Dil
            calismaSayfasi.Cells[3, 24] = "Y";  //Yabancı Dil
            calismaSayfasi.Cells[3, 25] = "Net";//Yabancı Dil


            int ogrenciSayisi = ogrenciXls.Count;
            progressBar1.Maximum = ogrenciSayisi;

            for (var i = 0; i < ogrenciSayisi; i++)
            {
                progressBar1.Value = i;
                label4.Text = $"Öğrenci sonuçları excele işleniyor {i + 1}/{ogrenciSayisi}";

                calismaSayfasi.Cells[4 + i, 1] = i + 1;
                calismaSayfasi.Cells[4 + i, 2] = ogrenciXls[i].IlceAdi;
                calismaSayfasi.Cells[4 + i, 3] = ogrenciXls[i].OkulAdi;
                calismaSayfasi.Cells[4 + i, 4] = ogrenciXls[i].AdiSoyadi;
                calismaSayfasi.Cells[4 + i, 5] = decimal.Round(ogrenciXls[i].SinavPuani, 3, MidpointRounding.AwayFromZero); //yüzdelik
                calismaSayfasi.Cells[4 + i, 6] = decimal.Round(ogrenciXls[i].YuzdelikDilim, 2, MidpointRounding.AwayFromZero); //yüzdelik
                calismaSayfasi.Cells[4 + i, 7] = ogrenciXls[i].ToplamNet;
                calismaSayfasi.Cells[4 + i, 8] = ogrenciXls[i].TurkceDogru;
                calismaSayfasi.Cells[4 + i, 9] = ogrenciXls[i].TurkceYanlis;
                calismaSayfasi.Cells[4 + i, 10] = decimal.Round(ogrenciXls[i].TurkceNet, 2, MidpointRounding.AwayFromZero); //yüzdelik
                calismaSayfasi.Cells[4 + i, 11] = ogrenciXls[i].MatDogru;
                calismaSayfasi.Cells[4 + i, 12] = ogrenciXls[i].MatYanlis;
                calismaSayfasi.Cells[4 + i, 13] = decimal.Round(ogrenciXls[i].MatNet, 2, MidpointRounding.AwayFromZero); //yüzdelik
                calismaSayfasi.Cells[4 + i, 14] = ogrenciXls[i].FenDogru;
                calismaSayfasi.Cells[4 + i, 15] = ogrenciXls[i].FenYanlis;
                calismaSayfasi.Cells[4 + i, 16] = decimal.Round(ogrenciXls[i].FenNet, 2, MidpointRounding.AwayFromZero); //yüzdelik
                calismaSayfasi.Cells[4 + i, 17] = ogrenciXls[i].InkDogru;
                calismaSayfasi.Cells[4 + i, 18] = ogrenciXls[i].InkYanlis;
                calismaSayfasi.Cells[4 + i, 19] = decimal.Round(ogrenciXls[i].InkNet, 2, MidpointRounding.AwayFromZero); //yüzdelik
                calismaSayfasi.Cells[4 + i, 20] = ogrenciXls[i].DinDogru;
                calismaSayfasi.Cells[4 + i, 21] = ogrenciXls[i].DinYanlis;
                calismaSayfasi.Cells[4 + i, 22] = decimal.Round(ogrenciXls[i].DinNet, 2, MidpointRounding.AwayFromZero); //yüzdelik
                calismaSayfasi.Cells[4 + i, 23] = ogrenciXls[i].IngDogru;
                calismaSayfasi.Cells[4 + i, 24] = ogrenciXls[i].IngYanlis;
                calismaSayfasi.Cells[4 + i, 25] = decimal.Round(ogrenciXls[i].IngNet, 2, MidpointRounding.AwayFromZero); //yüzdelik

            }
            progressBar1.Value = 0;

            //başlık2 exceldeki ikinci satır net doğru yanlış bilgilerinin olduğu satır
            int satirGenisligi = 25;

            Range baslik2 = calismaSayfasi.Range[calismaSayfasi.Cells[2, 1], calismaSayfasi.Cells[3, satirGenisligi]];
            baslik2.EntireRow.Font.Bold = true; //bold yap
            baslik2.Font.Size = 10;
            baslik2.Style.HorizontalAlignment = XlHAlign.xlHAlignCenter; //hücreyi yatay ortala
            baslik2.Style.VerticalAlignment = XlHAlign.xlHAlignCenter; //hücreyi dikey ortala
            baslik2.Cells.WrapText = true; //Metni kaydır
            baslik2.Borders.LineStyle = XlLineStyle.xlContinuous;

            Range veriler = calismaSayfasi.Range[calismaSayfasi.Cells[4, 1], calismaSayfasi.Cells[ogrenciSayisi + 3, satirGenisligi]];
            veriler.Font.Size = 9;
            veriler.Style.HorizontalAlignment = XlHAlign.xlHAlignCenter; //hücreyi yatay ortala
            veriler.Style.VerticalAlignment = XlHAlign.xlHAlignCenter; //hücreyi dikey ortala
            veriler.Borders.LineStyle = XlLineStyle.xlContinuous;


            //Range toplamNet = calismaSayfasi.Range[calismaSayfasi.Cells[3, 12], calismaSayfasi.Cells[ogrenciSayisi + 2, 12]];
            //toplamNet.NumberFormat = "#,###,###.00";

            //Range yuzdelikPuanToplam = calismaSayfasi.Range[calismaSayfasi.Cells[3, 10], calismaSayfasi.Cells[ogrenciSayisi + 2, 11]];
            //yuzdelikPuanToplam.NumberFormat = "#,###,###.000";

        }
        private void ExcelOkulSayfasi(List<LgsSonuc> okulXls, Workbook calismaKitabi)
        {
            Sheets xlSheets = calismaKitabi.Sheets;
            var calismaSayfasi = (Worksheet)xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);

            calismaSayfasi.Name = "OKUL";


            ExcelUtil.HucreBirlesitir(calismaSayfasi, 1, 1, 1, 24, 60); //başlık
            calismaSayfasi.Cells[1, 1] = sinavAdi;

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 7, 2, 9);
            calismaSayfasi.Cells[2, 7] = "Türkçe";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 10, 2, 12);
            calismaSayfasi.Cells[2, 10] = "Matematik";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 13, 2, 15);
            calismaSayfasi.Cells[2, 13] = "Fen Bilimleri";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 16, 2, 18);
            calismaSayfasi.Cells[2, 16] = "İnkılap Tarihi";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 19, 2, 21);
            calismaSayfasi.Cells[2, 19] = "DKAB";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 22, 2, 24);
            calismaSayfasi.Cells[2, 22] = "Yabancı Dil";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 1, 3, 1);
            calismaSayfasi.Cells[2, 1] = "No";


            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 2, 3, 2);
            calismaSayfasi.Cells[2, 2] = "İlçe Adı";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 3, 3, 3);
            calismaSayfasi.Cells[2, 3] = "Kurum Adı";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 4, 3, 4);
            calismaSayfasi.Cells[2, 4] = "Sınava Giren Öğr. S.";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 5, 3, 5);
            calismaSayfasi.Cells[2, 5] = "Sınav Puanı Ort.";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 6, 3, 6);
            calismaSayfasi.Cells[2, 6] = "Toplam Net Ort.";

            calismaSayfasi.Cells[3, 7] = "D";//Türkçe
            calismaSayfasi.Cells[3, 8] = "Y";//Türkçe
            calismaSayfasi.Cells[3, 9] = "Net";//Türkçe
            calismaSayfasi.Cells[3, 10] = "D";//Matematik
            calismaSayfasi.Cells[3, 11] = "Y";//Matematik
            calismaSayfasi.Cells[3, 12] = "Net";//Matematik
            calismaSayfasi.Cells[3, 13] = "D";//Fen Bilimleri
            calismaSayfasi.Cells[3, 14] = "Y";//Fen Bilimleri
            calismaSayfasi.Cells[3, 15] = "Net";//Fen Bilimleri
            calismaSayfasi.Cells[3, 16] = "D";//İnkılap Tarihi
            calismaSayfasi.Cells[3, 17] = "Y";//İnkılap Tarihi
            calismaSayfasi.Cells[3, 18] = "Net";//İnkılap Tarihi
            calismaSayfasi.Cells[3, 19] = "D";  //DKAB
            calismaSayfasi.Cells[3, 20] = "Y";  //DKAB
            calismaSayfasi.Cells[3, 21] = "Net";//DKAB
            calismaSayfasi.Cells[3, 22] = "D";  //Yabancı Dil
            calismaSayfasi.Cells[3, 23] = "Y";  //Yabancı Dil
            calismaSayfasi.Cells[3, 24] = "Net";//Yabancı Dil

            progressBar1.Maximum = okulXls.Count;
            for (var i = 0; i < okulXls.Count; i++)
            {
                label4.Text = $"Okul sonuçları excele işleniyor. {i + 1}/{okulXls.Count}";
                progressBar1.Value = i;
                calismaSayfasi.Cells[4 + i, 1] = i + 1;
                calismaSayfasi.Cells[4 + i, 2] = okulXls[i].IlceAdi;
                calismaSayfasi.Cells[4 + i, 3] = okulXls[i].OkulAdi;
                calismaSayfasi.Cells[4 + i, 4] = okulXls[i].OgrenciSayisi;
                calismaSayfasi.Cells[4 + i, 5] = decimal.Round(okulXls[i].SinavPuani, 3, MidpointRounding.AwayFromZero);
                calismaSayfasi.Cells[4 + i, 6] = decimal.Round(okulXls[i].ToplamNet, 2, MidpointRounding.AwayFromZero);
                calismaSayfasi.Cells[4 + i, 7] = decimal.Round(okulXls[i].TurkceDogru / (decimal)okulXls[i].OgrenciSayisi, 2, MidpointRounding.AwayFromZero);
                calismaSayfasi.Cells[4 + i, 8] = decimal.Round(okulXls[i].TurkceYanlis / (decimal)okulXls[i].OgrenciSayisi, 2, MidpointRounding.AwayFromZero);
                calismaSayfasi.Cells[4 + i, 9] = decimal.Round(okulXls[i].TurkceNet, 2, MidpointRounding.AwayFromZero); //yüzdelik
                calismaSayfasi.Cells[4 + i, 10] = decimal.Round(okulXls[i].MatDogru / (decimal)okulXls[i].OgrenciSayisi, 2, MidpointRounding.AwayFromZero);
                calismaSayfasi.Cells[4 + i, 11] = decimal.Round(okulXls[i].MatYanlis / (decimal)okulXls[i].OgrenciSayisi, 2, MidpointRounding.AwayFromZero);
                calismaSayfasi.Cells[4 + i, 12] = decimal.Round(okulXls[i].MatNet, 2, MidpointRounding.AwayFromZero); //yüzdelik
                calismaSayfasi.Cells[4 + i, 13] = decimal.Round(okulXls[i].FenDogru / (decimal)okulXls[i].OgrenciSayisi, 2, MidpointRounding.AwayFromZero);
                calismaSayfasi.Cells[4 + i, 14] = decimal.Round(okulXls[i].FenYanlis / (decimal)okulXls[i].OgrenciSayisi, 2, MidpointRounding.AwayFromZero);
                calismaSayfasi.Cells[4 + i, 15] = decimal.Round(okulXls[i].FenNet, 2, MidpointRounding.AwayFromZero); //yüzdelik
                calismaSayfasi.Cells[4 + i, 16] = decimal.Round(okulXls[i].InkDogru / (decimal)okulXls[i].OgrenciSayisi, 2, MidpointRounding.AwayFromZero);
                calismaSayfasi.Cells[4 + i, 17] = decimal.Round(okulXls[i].InkYanlis / (decimal)okulXls[i].OgrenciSayisi, 2, MidpointRounding.AwayFromZero);
                calismaSayfasi.Cells[4 + i, 18] = decimal.Round(okulXls[i].InkNet, 2, MidpointRounding.AwayFromZero); //yüzdelik
                calismaSayfasi.Cells[4 + i, 19] = decimal.Round(okulXls[i].DinDogru / (decimal)okulXls[i].OgrenciSayisi, 2, MidpointRounding.AwayFromZero);
                calismaSayfasi.Cells[4 + i, 20] = decimal.Round(okulXls[i].DinYanlis / (decimal)okulXls[i].OgrenciSayisi, 2, MidpointRounding.AwayFromZero);
                calismaSayfasi.Cells[4 + i, 21] = decimal.Round(okulXls[i].DinNet, 2, MidpointRounding.AwayFromZero); //yüzdelik
                calismaSayfasi.Cells[4 + i, 22] = decimal.Round(okulXls[i].IngDogru / (decimal)okulXls[i].OgrenciSayisi, 2, MidpointRounding.AwayFromZero);
                calismaSayfasi.Cells[4 + i, 23] = decimal.Round(okulXls[i].IngYanlis / (decimal)okulXls[i].OgrenciSayisi, 2, MidpointRounding.AwayFromZero);
                calismaSayfasi.Cells[4 + i, 24] = decimal.Round(okulXls[i].IngNet, 2, MidpointRounding.AwayFromZero); //yüzdelik

            }
            progressBar1.Value = 0;

            //başlık2 exceldeki ikinci satır net doğru yanlış bilgilerinin olduğu satır
            int satirGenisligi = 24;

            Range baslik2 = calismaSayfasi.Range[calismaSayfasi.Cells[2, 1], calismaSayfasi.Cells[3, satirGenisligi]];
            baslik2.EntireRow.Font.Bold = true; //bold yap
            baslik2.Font.Size = 10;
            baslik2.Style.HorizontalAlignment = XlHAlign.xlHAlignCenter; //hücreyi yatay ortala
            baslik2.Style.VerticalAlignment = XlHAlign.xlHAlignCenter; //hücreyi dikey ortala
            baslik2.Cells.WrapText = true; //Metni kaydır
            baslik2.Borders.LineStyle = XlLineStyle.xlContinuous;

            Range veriler = calismaSayfasi.Range[calismaSayfasi.Cells[4, 1], calismaSayfasi.Cells[okulXls.Count + 3, satirGenisligi]];
            veriler.Font.Size = 9;
            veriler.Style.HorizontalAlignment = XlHAlign.xlHAlignCenter; //hücreyi yatay ortala
            veriler.Style.VerticalAlignment = XlHAlign.xlHAlignCenter; //hücreyi dikey ortala
            veriler.Borders.LineStyle = XlLineStyle.xlContinuous;



        }
        private void ExcelIlceSayfasi(List<LgsSonuc> ilceXls, Workbook calismaKitabi)
        {
            Sheets xlSheets = calismaKitabi.Sheets;
            var calismaSayfasi = (Worksheet)xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);

            calismaSayfasi.Name = "İLÇE";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 1, 1, 1, 23, 60); //başlık
            calismaSayfasi.Cells[1, 1] = sinavAdi;

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 1, 3, 1);
            calismaSayfasi.Cells[2, 1] = "No";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 2, 3, 2);
            calismaSayfasi.Cells[2, 2] = "İlçe Adı";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 3, 3, 3);
            calismaSayfasi.Cells[2, 3] = "Sınava Giren Öğr. S.";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 4, 3, 4);
            calismaSayfasi.Cells[2, 4] = "Sınav Puanı Ort.";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 5, 3, 5);
            calismaSayfasi.Cells[2, 5] = "Toplam Net Ort.";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 6, 2, 8);
            calismaSayfasi.Cells[2, 6] = "Türkçe";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 9, 2, 11);
            calismaSayfasi.Cells[2, 9] = "Matematik";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 12, 2, 14);
            calismaSayfasi.Cells[2, 12] = "Fen Bilimleri";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 15, 2, 17);
            calismaSayfasi.Cells[2, 15] = "İnkılap Tarihi";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 18, 2, 20);
            calismaSayfasi.Cells[2, 18] = "DKAB";

            ExcelUtil.HucreBirlesitir(calismaSayfasi, 2, 21, 2, 23);
            calismaSayfasi.Cells[2, 21] = "Yabancı Dil";

            calismaSayfasi.Cells[3, 6] = "D";//Türkçe
            calismaSayfasi.Cells[3, 7] = "Y";//Türkçe
            calismaSayfasi.Cells[3, 8] = "Net";//Türkçe
            calismaSayfasi.Cells[3, 9] = "D";//Matematik
            calismaSayfasi.Cells[3, 10] = "Y";//Matematik
            calismaSayfasi.Cells[3, 11] = "Net";//Matematik
            calismaSayfasi.Cells[3, 12] = "D";//Fen Bilimleri
            calismaSayfasi.Cells[3, 13] = "Y";//Fen Bilimleri
            calismaSayfasi.Cells[3, 14] = "Net";//Fen Bilimleri
            calismaSayfasi.Cells[3, 15] = "D";//İnkılap Tarihi
            calismaSayfasi.Cells[3, 16] = "Y";//İnkılap Tarihi
            calismaSayfasi.Cells[3, 17] = "Net";//İnkılap Tarihi
            calismaSayfasi.Cells[3, 18] = "D";  //DKAB
            calismaSayfasi.Cells[3, 19] = "Y";  //DKAB
            calismaSayfasi.Cells[3, 20] = "Net";//DKAB
            calismaSayfasi.Cells[3, 21] = "D";  //Yabancı Dil
            calismaSayfasi.Cells[3, 22] = "Y";  //Yabancı Dil
            calismaSayfasi.Cells[3, 23] = "Net";//Yabancı Dil

            progressBar1.Maximum = ilceXls.Count;
            for (var i = 0; i < ilceXls.Count; i++)
            {
                label4.Text = $"İlçe sonuçları excele işleniyor {i + 1}/{ilceXls.Count}";
                progressBar1.Value = i;

                calismaSayfasi.Cells[4 + i, 1] = i + 1;
                calismaSayfasi.Cells[4 + i, 2] = ilceXls[i].IlceAdi;
                calismaSayfasi.Cells[4 + i, 3] = ilceXls[i].OgrenciSayisi;
                calismaSayfasi.Cells[4 + i, 4] = decimal.Round(ilceXls[i].SinavPuani, 3, MidpointRounding.AwayFromZero);
                calismaSayfasi.Cells[4 + i, 5] = decimal.Round(ilceXls[i].ToplamNet, 2, MidpointRounding.AwayFromZero);
                calismaSayfasi.Cells[4 + i, 6] = decimal.Round(ilceXls[i].TurkceDogru / (decimal)ilceXls[i].OgrenciSayisi, 2, MidpointRounding.AwayFromZero);
                calismaSayfasi.Cells[4 + i, 7] = decimal.Round(ilceXls[i].TurkceYanlis / (decimal)ilceXls[i].OgrenciSayisi, 2, MidpointRounding.AwayFromZero);
                calismaSayfasi.Cells[4 + i, 8] = decimal.Round(ilceXls[i].TurkceNet, 2, MidpointRounding.AwayFromZero); //yüzdelik
                calismaSayfasi.Cells[4 + i, 9] = decimal.Round(ilceXls[i].MatDogru / (decimal)ilceXls[i].OgrenciSayisi, 2, MidpointRounding.AwayFromZero);
                calismaSayfasi.Cells[4 + i, 10] = decimal.Round(ilceXls[i].MatYanlis / (decimal)ilceXls[i].OgrenciSayisi, 2, MidpointRounding.AwayFromZero);
                calismaSayfasi.Cells[4 + i, 11] = decimal.Round(ilceXls[i].MatNet, 2, MidpointRounding.AwayFromZero); //yüzdelik
                calismaSayfasi.Cells[4 + i, 12] = decimal.Round(ilceXls[i].FenDogru / (decimal)ilceXls[i].OgrenciSayisi, 2, MidpointRounding.AwayFromZero);
                calismaSayfasi.Cells[4 + i, 13] = decimal.Round(ilceXls[i].FenYanlis / (decimal)ilceXls[i].OgrenciSayisi, 2, MidpointRounding.AwayFromZero);
                calismaSayfasi.Cells[4 + i, 14] = decimal.Round(ilceXls[i].FenNet, 2, MidpointRounding.AwayFromZero); //yüzdelik
                calismaSayfasi.Cells[4 + i, 15] = decimal.Round(ilceXls[i].InkDogru / (decimal)ilceXls[i].OgrenciSayisi, 2, MidpointRounding.AwayFromZero);
                calismaSayfasi.Cells[4 + i, 16] = decimal.Round(ilceXls[i].InkYanlis / (decimal)ilceXls[i].OgrenciSayisi, 2, MidpointRounding.AwayFromZero);
                calismaSayfasi.Cells[4 + i, 17] = decimal.Round(ilceXls[i].InkNet, 2, MidpointRounding.AwayFromZero); //yüzdelik
                calismaSayfasi.Cells[4 + i, 18] = decimal.Round(ilceXls[i].DinDogru / (decimal)ilceXls[i].OgrenciSayisi, 2, MidpointRounding.AwayFromZero);
                calismaSayfasi.Cells[4 + i, 19] = decimal.Round(ilceXls[i].DinYanlis / (decimal)ilceXls[i].OgrenciSayisi, 2, MidpointRounding.AwayFromZero);
                calismaSayfasi.Cells[4 + i, 20] = decimal.Round(ilceXls[i].DinNet, 2, MidpointRounding.AwayFromZero); //yüzdelik
                calismaSayfasi.Cells[4 + i, 21] = decimal.Round(ilceXls[i].IngDogru / (decimal)ilceXls[i].OgrenciSayisi, 2, MidpointRounding.AwayFromZero);
                calismaSayfasi.Cells[4 + i, 22] = decimal.Round(ilceXls[i].IngYanlis / (decimal)ilceXls[i].OgrenciSayisi, 2, MidpointRounding.AwayFromZero);
                calismaSayfasi.Cells[4 + i, 23] = decimal.Round(ilceXls[i].IngNet, 2, MidpointRounding.AwayFromZero); //yüzdelik
            }

            progressBar1.Value = 0;

            //başlık2 exceldeki ikinci satır net doğru yanlış bilgilerinin olduğu satır
            int satirGenisligi = 23;

            Range baslik2 = calismaSayfasi.Range[calismaSayfasi.Cells[1, 1], calismaSayfasi.Cells[3, satirGenisligi]];
            baslik2.EntireRow.Font.Bold = true; //bold yap
            baslik2.Font.Size = 10;
            baslik2.Style.HorizontalAlignment = XlHAlign.xlHAlignCenter; //hücreyi yatay ortala
            baslik2.Style.VerticalAlignment = XlHAlign.xlHAlignCenter; //hücreyi dikey ortala
            baslik2.Cells.WrapText = true; //Metni kaydır
            baslik2.Borders.LineStyle = XlLineStyle.xlContinuous;

            Range veriler = calismaSayfasi.Range[calismaSayfasi.Cells[4, 1], calismaSayfasi.Cells[ilceXls.Count + 3, satirGenisligi]];
            veriler.Font.Size = 9;
            veriler.Style.HorizontalAlignment = XlHAlign.xlHAlignCenter; //hücreyi yatay ortala
            veriler.Style.VerticalAlignment = XlHAlign.xlHAlignCenter; //hücreyi dikey ortala
            veriler.Borders.LineStyle = XlLineStyle.xlContinuous;


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

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("http://osmancelik.com.tr");
        }

        private void FormLgs_FormClosing(object sender, FormClosingEventArgs e)
        {
            FormGiris frm = new FormGiris();
            frm.Show();
        }

    }
}
