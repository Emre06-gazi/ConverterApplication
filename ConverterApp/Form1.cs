using System;
using System.IO;
using System.Windows.Forms;
using iTextSharp.text.pdf.parser;
using iTextSharp.text.pdf;
using Path = System.IO.Path;
using System.Drawing;
using System.Drawing.Imaging;
using ImageFormat = System.Drawing.Imaging.ImageFormat;
using Microsoft.Office.Interop.Word;
using Tesseract;
using Spire.Doc;
using Aspose.Words;
using Aspose.Words.Saving;
using System.Linq;

namespace ConverterPro
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            InitializeComboBoxes();
        }

        private void InitializeComboBoxes()
        {
            comboBox2.Items.AddRange(new string[] { ".pdf", ".docx", ".txt", ".jpg", ".png", ".ico", ".csv", ".html", ".pptx", ".xls" });
            comboBox3.Items.AddRange(new string[] { ".txt", ".docx", ".pdf" });
            comboBox5.Items.AddRange(new string[] { "Türkçe", "English", "Русский" });
        }

        private void SelectFile(TextBox textBox, ComboBox comboBox, string filter)
        {
            openFileDialog1.Filter = filter;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string selectedFilePath = openFileDialog1.FileName;
                textBox.Text = selectedFilePath;
                comboBox.Text = Path.GetExtension(selectedFilePath);
            }
        }

        // DOCX Dönüşümleri
        private string ConvertWordToText(string docxFilePath)
        {
            string text = "";
            try
            {
                var wordApp = new Microsoft.Office.Interop.Word.Application();
                var doc = wordApp.Documents.Open(docxFilePath);
                text = doc.Content.Text;
                doc.Close();
                wordApp.Quit();
            }
            catch (Exception ex)
            {
                HandleError("Word'den metne dönüştürme sırasında bir hata oluştu: " + ex.Message);
            }
            return text;
        }

        private string ConvertDocxToPdf(string docxFilePath, string pdfFilePath)
        {
            string text = "";
            try
            {
                var wordApp = new Microsoft.Office.Interop.Word.Application();
                var doc = wordApp.Documents.Open(docxFilePath);
                text = doc.Content.Text;
                doc.SaveAs2(pdfFilePath, WdSaveFormat.wdFormatPDF);
                doc.Close();
                wordApp.Quit();
            }
            catch (Exception ex)
            {
                HandleError("Word'den metne dönüştürme sırasında bir hata oluştu: " + ex.Message);
            }
            return text;
        }

        private void ConvertDocxToJpgPages(string docxFilePath, string jpgFolderPath)
        {
            try
            {
                Spire.Doc.Document doc = new Spire.Doc.Document();
                doc.LoadFromFile(docxFilePath);

                for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
                {
                    string jpgFileName = $"{Path.GetFileNameWithoutExtension(docxFilePath)}_page{pageIndex + 1}.jpg";
                    string jpgFilePath = Path.Combine(jpgFolderPath, jpgFileName);

                    using (System.Drawing.Image image = doc.SaveToImages(pageIndex, Spire.Doc.Documents.ImageType.Bitmap))
                    {
                        using (Bitmap bitmap = new Bitmap(image))
                        {
                            bitmap.Save(jpgFilePath, ImageFormat.Jpeg);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                HandleError("Docx'ten JPG'ye dönüşüm sırasında bir hata oluştu: " + ex.Message);
            }
        }

        // PDF Dönüşümleri
        private string ConvertPdfToText(string pdfFilePath)
        {
            string text = "";
            try
            {
                using (PdfReader reader = new PdfReader(pdfFilePath))
                {
                    for (int page = 1; page <= reader.NumberOfPages; page++)
                    {
                        text += PdfTextExtractor.GetTextFromPage(reader, page);
                    }
                }
            }
            catch (Exception ex)
            {
                HandleError("PDF'den metne dönüştürme sırasında bir hata oluştu: " + ex.Message);
            }
            return text;
        }

        private void ConvertPdfToDocx(string pdfFilePath, string docxFilePath)
        {
            try
            {
                Aspose.Words.Document pdfDocument = new Aspose.Words.Document(pdfFilePath);

                pdfDocument.Save(docxFilePath, SaveFormat.Docx);
            }
            catch (Exception ex)
            {
                HandleError("PDF'den DOCX'e dönüşüm sırasında bir hata oluştu: " + ex.Message);
            }
        }

        private void ConvertPdfToJpgPages(string pdfFilePath, string jpgFolderPath, int targetDpi)
        {
            try
            {
                using (PdfReader reader = new PdfReader(pdfFilePath))
                {
                    for (int page = 1; page <= reader.NumberOfPages; page++)
                    {
                        string jpgFileName = $"{Path.GetFileNameWithoutExtension(pdfFilePath)}_page{page}.jpg";
                        string jpgFilePath = Path.Combine(jpgFolderPath, jpgFileName);

                        Bitmap bmp;

                        // Hedef boyutları ayarla (1428x2020 piksel)
                        float pageWidth = 1428;
                        float pageHeight = 2020;

                        // Hedef DPI'ı kullanarak boyutları ayarla
                        bmp = new Bitmap((int)pageWidth, (int)pageHeight);
                        bmp.SetResolution(targetDpi, targetDpi);

                        using (Graphics graphics = Graphics.FromImage(bmp))
                        {
                            graphics.Clear(Color.White);
                            graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;

                            // PDF sayfasını JPG'ye dönüştür
                            string text = PdfTextExtractor.GetTextFromPage(reader, page);
                            string[] lines = text.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

                            using (SolidBrush brush = new SolidBrush(Color.Black))
                            {
                                for (int i = 0; i < lines.Length; i++)
                                {
                                    string line = lines[i];
                                    PointF drawPoint = new PointF(10, i * 20);
                                    graphics.DrawString(line, Font, brush, drawPoint);
                                }
                            }
                        }

                        // JPG dosyasını kaydet
                        bmp.Save(jpgFilePath, ImageFormat.Jpeg);
                        bmp.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                HandleError("PDF'den JPG'ye dönüşüm sırasında bir hata oluştu: " + ex.Message);
            }
        }

        // TXT Dönüşümleri
        private void ConvertTxtToJpg(string txtFilePath, string jpgFolderPath)
        {
            try
            {
                string[] lines = File.ReadAllLines(txtFilePath);
                int lineHeight = 40; // Metin satır aralığı (ayarlamalar yapabilirsiniz)
                int padding = 10; // Kenar boşlukları (ayarlamalar yapabilirsiniz)

                int imageHeight = 1428;
                int imageWidth = 2020;

                Bitmap bmp = new Bitmap(imageWidth, imageHeight);

                using (Graphics graphics = Graphics.FromImage(bmp))
                {
                    graphics.Clear(Color.White); // Arkaplanı beyaz olarak temizle
                    graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias; // Metin kalitesini iyileştir

                    using (SolidBrush brush = new SolidBrush(Color.Black))
                    {
                        for (int i = 0; i < lines.Length; i++)
                        {
                            string line = lines[i];
                            SizeF textSize = graphics.MeasureString(line, Font);
                            PointF drawPoint = new PointF(padding, i * lineHeight + padding);

                            // Metin satırını boyutlandırarak sığdır
                            if (textSize.Height + drawPoint.Y > imageHeight)
                                break;

                            graphics.DrawString(line, Font, brush, drawPoint);
                        }
                    }
                }

                string jpgFilePath = Path.Combine(jpgFolderPath, Path.GetFileNameWithoutExtension(txtFilePath) + ".jpg");
                bmp.Save(jpgFilePath, ImageFormat.Jpeg);
            }
            catch (Exception ex)
            {
                HandleError("TXT'den JPG'ye dönüşüm sırasında bir hata oluştu: " + ex.Message);
            }
        }


        private void ConvertTxtToPng(string txtFilePath, string pngFolderPath)
        {
            try
            {
                string[] lines = File.ReadAllLines(txtFilePath);
                int lineHeight = 20; // Metin satır aralığı (ayarlamalar yapabilirsiniz)
                int padding = 10; // Kenar boşlukları (ayarlamalar yapabilirsiniz)

                int imageHeight = 1428;
                int imageWidth = 2020;

                using (Bitmap bmp = new Bitmap(imageWidth, imageHeight))
                {
                    using (Graphics graphics = Graphics.FromImage(bmp))
                    {
                        graphics.Clear(Color.White); // Arkaplanı beyaz olarak temizle
                        graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias; // Metin kalitesini iyileştir

                        using (SolidBrush brush = new SolidBrush(Color.Black))
                        {
                            for (int i = 0; i < lines.Length; i++)
                            {
                                string line = lines[i];
                                PointF drawPoint = new PointF(padding, i * lineHeight + padding);
                                graphics.DrawString(line, Font, brush, drawPoint);
                            }
                        }
                    }

                    string pngFilePath = Path.Combine(pngFolderPath, Path.GetFileNameWithoutExtension(txtFilePath) + ".png");
                    bmp.Save(pngFilePath, ImageFormat.Png);
                }
            }
            catch (Exception ex)
            {
                HandleError("TXT'den PNG'ye dönüşüm sırasında bir hata oluştu: " + ex.Message);
            }
        }

        private void ConvertTxtToDocx(string txtFilePath, string docxFilePath)
        {
            try
            {
                string[] lines = File.ReadAllLines(txtFilePath);
                string text = string.Join(Environment.NewLine, lines);

                var wordApp = new Microsoft.Office.Interop.Word.Application();
                var doc = wordApp.Documents.Add();
                doc.Content.Text = text;

                doc.SaveAs2(docxFilePath, WdSaveFormat.wdFormatDocumentDefault);
                doc.Close(false);
                wordApp.Quit();
            }
            catch (Exception ex)
            {
                HandleError("TXT'den DOCX'e dönüşüm sırasında bir hata oluştu: " + ex.Message);
            }
        }

        // JPG Dönüşümleri
        private void ConvertJpgToPng(string jpgFilePath, string pngFolderPath)
        {
            try
            {
                using (Bitmap jpgImage = new Bitmap(jpgFilePath))
                {
                    int maxWidth = 7000; // Resim maksimum genişlik

                    int originalWidth = jpgImage.Width;
                    int originalHeight = jpgImage.Height;

                    int imageWidth = originalWidth <= maxWidth ? originalWidth : maxWidth;
                    int imageHeight = (int)((double)originalHeight / originalWidth * imageWidth);

                    using (Bitmap resizedImage = new Bitmap(imageWidth, imageHeight))
                    {
                        using (Graphics graphics = Graphics.FromImage(resizedImage))
                        {
                            graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                            graphics.DrawImage(jpgImage, 0, 0, imageWidth, imageHeight);
                        }

                        string pngFilePath = Path.Combine(pngFolderPath, Path.GetFileNameWithoutExtension(jpgFilePath) + ".png");

                        // Resim formatını belirle
                        ImageFormat format = ImageFormat.Png;

                        // JPEG için uygun kodlayıcıyı belirle
                        ImageCodecInfo jpgEncoder = GetEncoderInfo(format);

                        // Kodlayıcı parametrelerini belirle
                        EncoderParameters encoderParameters = new EncoderParameters(1);
                        encoderParameters.Param[0] = new EncoderParameter(Encoder.Quality, 100L); // Kaliteyi ayarla (örneğin, 100 kalite)

                        // Resmi kaydet
                        resizedImage.Save(pngFilePath, jpgEncoder, encoderParameters);
                    }
                }
            }
            catch (Exception ex)
            {
                HandleError("JPG'den PNG'ye dönüşüm sırasında bir hata oluştu: " + ex.Message);
            }
        }

        private void ConvertJpgToIco(string jpgFilePath, string icoFilePath)
        {
            try
            {
                using (Bitmap jpgImage = new Bitmap(jpgFilePath))
                {
                    // İcon boyutunu ayarlayın
                    int iconWidth = 300;
                    int iconHeight = 300;

                    // İcon oluşturun
                    using (Bitmap iconImage = new Bitmap(iconWidth, iconHeight))
                    {
                        using (Graphics graphics = Graphics.FromImage(iconImage))
                        {
                            graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                            graphics.DrawImage(jpgImage, 0, 0, iconWidth, iconHeight);
                        }

                        // İcon'u ICO dosyasına kaydedin
                        using (FileStream iconStream = new FileStream(icoFilePath, FileMode.Create))
                        {
                            Icon icon = Icon.FromHandle(iconImage.GetHicon());
                            icon.Save(iconStream);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                HandleError("JPG'den ICO'ya dönüşüm sırasında bir hata oluştu: " + ex.Message);
            }
        }

        private void ConvertJpgToPdf(string[] jpgFilePaths, string pdfFilePath)
        {
            try
            {
                using (iTextSharp.text.Document pdfDoc = new iTextSharp.text.Document())
                {
                    PdfWriter.GetInstance(pdfDoc, new FileStream(pdfFilePath, FileMode.Create));

                    pdfDoc.Open();

                    foreach (string jpgFilePath in jpgFilePaths)
                    {
                        using (FileStream jpgStream = new FileStream(jpgFilePath, FileMode.Open, FileAccess.Read))
                        {
                            iTextSharp.text.Image jpgImage = iTextSharp.text.Image.GetInstance(jpgStream);

                            // JPG resminin boyutunu belirleyin (örneğin, A4 kağıdına sığacak şekilde)
                            jpgImage.ScaleToFit(pdfDoc.PageSize.Width - 60, pdfDoc.PageSize.Height - 60);

                            // Resmi 20 birim sağdan, alttan, soldan ve üstten ekleyin
                            jpgImage.SetAbsolutePosition(25, 25);

                            // Kalite ayarlarını yapın (örneğin, 100 kalite)
                            jpgImage.CompressionLevel = 100;

                            // JPG resimlerini PDF dosyasına ekleyin
                            pdfDoc.Add(jpgImage);
                        }
                    }

                    pdfDoc.Close();
                }
            }
            catch (Exception ex)
            {
                HandleError("JPG'den PDF'ye dönüşüm sırasında bir hata oluştu: " + ex.Message);
            }
        }

        private string ConvertJpgToTxt(string jpgFilePath)
        {
            string text = "";
            try
            {
                using (var engine = new TesseractEngine(@"C:\Users\emre_\OneDrive\Masaüstü\ConverterApplication\ConverterPro\tessdata", "eng", EngineMode.Default))
                {
                    using (var img = Pix.LoadFromFile(jpgFilePath))
                    {
                        using (var page = engine.Process(img))
                        {
                            text = page.GetText();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                HandleError("JPG'den metne dönüştürme sırasında bir hata oluştu: " + ex.Message);
            }
            return text;
        }

        private void ConvertJpgToDocx(string jpgFilePath, string docxFilePath)
        {
            try
            {
                // Tesseract OCR motorunu başlatma ve dil verisini yükleme
                using (var engine = new TesseractEngine(@"C:\Users\emre_\OneDrive\Masaüstü\ConverterApplication\ConverterPro\tessdata", "eng", EngineMode.Default))
                {
                    using (var image = Pix.LoadFromFile(jpgFilePath))
                    {
                        using (var page = engine.Process(image))
                        {
                            // OCR ile çevrilen metni al
                            string translatedText = page.GetText();

                            // Spire.Doc belgesi oluşturma ve bir bölüme metni ekleme
                            Spire.Doc.Document doc = new Spire.Doc.Document();
                            Spire.Doc.Section section = doc.AddSection();
                            Spire.Doc.Documents.Paragraph paragraph = section.AddParagraph();
                            paragraph.AppendText(translatedText);

                            // DOCX dosyasını kaydetme
                            doc.SaveToFile(docxFilePath, FileFormat.Docx);
                            doc.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                HandleError("JPG'den DOCX'e dönüşüm sırasında bir hata oluştu: " + ex.Message);
            }
        }

        // Dil Paketleri İçin Resimden - Metine Dönüşüm
        private string ConvertJpgToTxtLanguage(string imagePath, string selectedLanguage)
        {
            string text = "";
            try
            {
                string languageCode = GetTesseractLanguageCode(selectedLanguage);
                using (var engine = new TesseractEngine(@"C:\Users\emre_\OneDrive\Masaüstü\ConverterApplication\ConverterPro\tessdata", languageCode, EngineMode.Default))
                {
                    using (var img = Pix.LoadFromFile(imagePath))
                    {
                        using (var page = engine.Process(img))
                        {
                            text = page.GetText();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                HandleError("Metni çevirme sırasında bir hata oluştu: " + ex.Message);
            }
            return text;
        }

        private void ConvertJpgToDocxLanguage(string jpgFilePath, string docxFilePath, string selectedLanguage)
        {
            try
            {
                // Tesseract OCR motorunu başlatma ve dil verisini yükleme
                string languageCode = GetTesseractLanguageCode(selectedLanguage);
                using (var engine = new TesseractEngine(@"C:\Users\emre_\OneDrive\Masaüstü\ConverterApplication\ConverterPro\tessdata", languageCode, EngineMode.Default))
                {
                    using (var image = Pix.LoadFromFile(jpgFilePath))
                    {
                        using (var page = engine.Process(image))
                        {
                            // OCR ile çevrilen metni al
                            string translatedText = page.GetText();

                            // Spire.Doc belgesi oluşturma ve bir bölüme metni ekleme
                            Spire.Doc.Document doc = new Spire.Doc.Document();
                            Spire.Doc.Section section = doc.AddSection();
                            Spire.Doc.Documents.Paragraph paragraph = section.AddParagraph();
                            paragraph.AppendText(translatedText);

                            // DOCX dosyasını kaydetme
                            doc.SaveToFile(docxFilePath, FileFormat.Docx);
                            doc.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                HandleError("JPG'den DOCX'e dönüşüm sırasında bir hata oluştu: " + ex.Message);
            }
        }


        // PNG Dönüşümleri
        private void ConvertPngToJpg(string pngFilePath, string jpgFolderPath)
        {
            try
            {
                using (Image pngImage = Image.FromFile(pngFilePath))
                {
                    // Yüksek kaliteli JPG oluşturmak için EncoderParameters kullanabilirsiniz
                    System.Drawing.Imaging.EncoderParameters encoderParameters = new System.Drawing.Imaging.EncoderParameters(1);
                    encoderParameters.Param[0] = new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 100L); // Kaliteyi ayarlayın (örneğin, 100 kalite)

                    ImageCodecInfo jpgCodec = GetEncoderInfo(ImageFormat.Jpeg);

                    string jpgFilePath = Path.Combine(jpgFolderPath, Path.GetFileNameWithoutExtension(pngFilePath) + ".jpg");
                    pngImage.Save(jpgFilePath, jpgCodec, encoderParameters);

                }
            }
            catch (Exception ex)
            {
                HandleError("PNG'den JPG'ye dönüşüm sırasında bir hata oluştu: " + ex.Message);
            }
        }

        private void ConvertPngToIco(string pngFilePath, string icoFilePath)
        {
            try
            {
                using (FileStream pngStream = new FileStream(pngFilePath, FileMode.Open))
                {
                    using (Image pngImage = Image.FromStream(pngStream))
                    {
                        // Icon boyutunu ayarlayın (genellikle 16x16 veya 32x32)
                        int iconWidth = 128;
                        int iconHeight = 128;

                        // Icon oluşturun
                        using (Bitmap iconImage = new Bitmap(iconWidth, iconHeight))
                        {
                            using (Graphics graphics = Graphics.FromImage(iconImage))
                            {
                                graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                graphics.DrawImage(pngImage, 0, 0, iconWidth, iconHeight);
                            }

                            // Icon'u ICO dosyasına kaydedin
                            using (FileStream icoStream = new FileStream(icoFilePath, FileMode.Create))
                            {
                                iconImage.Save(icoStream, ImageFormat.Bmp);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                HandleError("PNG'den ICO'ya dönüşüm sırasında bir hata oluştu: " + ex.Message);
            }
        }

        private ImageCodecInfo GetEncoderInfo(ImageFormat format)
        {
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageDecoders();
            foreach (ImageCodecInfo codec in codecs)
            {
                if (codec.FormatID == format.Guid)
                {
                    return codec;
                }
            }
            return null;
        }

        private void ConvertPngToPdf(string[] pngFilePaths, string pdfFilePath)
        {
            try
            {
                using (iTextSharp.text.Document pdfDoc = new iTextSharp.text.Document())
                {
                    PdfWriter.GetInstance(pdfDoc, new FileStream(pdfFilePath, FileMode.Create));
                    pdfDoc.Open();

                    foreach (string pngFilePath in pngFilePaths)
                    {
                        using (FileStream pngStream = new FileStream(pngFilePath, FileMode.Open, FileAccess.Read))
                        {
                            iTextSharp.text.Image pngImage = iTextSharp.text.Image.GetInstance(pngStream);

                            pngImage.ScaleToFit(pdfDoc.PageSize.Width - 60, pdfDoc.PageSize.Height - 60);

                            pngImage.SetAbsolutePosition(25, 25);

                            // PNG resimlerini PDF dosyasına ekle
                            pdfDoc.Add(pngImage);
                        }
                    }

                    pdfDoc.Close();
                }
            }
            catch (Exception ex)
            {
                HandleError("PNG'den PDF'ye dönüşüm sırasında bir hata oluştu: " + ex.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SelectFile(textBox1, comboBox1, "Tüm Dosyalar|*.*");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SelectFile(textBox2, comboBox4, "Resim Dosyaları|*.jpg;*.png;*.bmp|Tüm Dosyalar|*.*");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string selectedFilePath = openFileDialog1.FileName;
            string selectedExtension = Path.GetExtension(selectedFilePath);

            if (string.IsNullOrEmpty(selectedFilePath) || string.IsNullOrEmpty(selectedExtension))
            {
                MessageBox.Show("Lütfen bir dosya seçiniz!");
                return;
            }

            string targetExtension = comboBox2.SelectedItem?.ToString();

            if (targetExtension != null)
            {
                try
                {
                    switch (selectedExtension)
                    {
                        case ".pdf":
                            HandlePdfConversion(selectedFilePath, targetExtension);
                            break;
                        case ".docx":
                            HandleDocxConversion(selectedFilePath, targetExtension);
                            break;
                        case ".txt":
                            HandleTxtConversion(selectedFilePath, targetExtension);
                            break;
                        case ".jpg":
                            HandleJpgConversion(selectedFilePath, targetExtension);
                            break;
                        case ".png":
                            HandlePngConversion(selectedFilePath, targetExtension);
                            break;
                        default:
                            MessageBox.Show("Bu uzantı için dönüşüm henüz desteklenmiyor.");
                            break;
                    }
                }
                catch (Exception ex)
                {
                    HandleError("Dönüşüm sırasında bir hata oluştu: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Lütfen bir hedef uzantı seçiniz!");
            }
        }

        private void HandlePdfConversion(string pdfFilePath, string targetExtension)
        {
            string convertedText = ConvertPdfToText(pdfFilePath);

            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(pdfFilePath);
            string newFileName = fileNameWithoutExtension + targetExtension;

            saveFileDialog1.FileName = newFileName;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string saveFilePath = saveFileDialog1.FileName;

                switch (targetExtension)
                {
                    case ".txt":
                        File.WriteAllText(saveFilePath, convertedText);
                        MessageBox.Show("Dosya başarıyla kaydedildi.");
                        break;
                    case ".docx":
                        ConvertPdfToDocx(pdfFilePath, saveFilePath);
                        MessageBox.Show("PDF başarıyla DOCX'e dönüştürüldü.");
                        break;
                    case ".jpg":
                        string jpgFolderPath = Path.GetDirectoryName(pdfFilePath);
                        int targetDpi = 300;
                        ConvertPdfToJpgPages(pdfFilePath, jpgFolderPath, targetDpi);
                        MessageBox.Show("PDF başarıyla JPG'lere dönüştürüldü.");
                        break;
                    default:
                        MessageBox.Show("Bu hedef format için dönüşüm henüz desteklenmiyor.");
                        break;
                }
            }
        }

        private void HandleDocxConversion(string docxFilePath, string targetExtension)
        {
            string convertedText = ConvertWordToText(docxFilePath);

            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(docxFilePath);
            string newFileName = fileNameWithoutExtension + targetExtension;

            saveFileDialog1.FileName = newFileName;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string saveFilePath = saveFileDialog1.FileName;

                switch (targetExtension)
                {
                    case ".txt":
                        File.WriteAllText(saveFilePath, convertedText);
                        MessageBox.Show("Dosya başarıyla kaydedildi.");
                        break;
                    case ".pdf":
                        convertedText = ConvertDocxToPdf(docxFilePath, saveFilePath);
                        MessageBox.Show("DOCX başarıyla PDF'ye dönüştürüldü.");
                        break;
                    case ".jpg":
                        ConvertDocxToJpgPages(docxFilePath, Path.GetDirectoryName(saveFilePath));
                        MessageBox.Show("DOCX başarıyla JPG'lere dönüştürüldü.");
                        break;
                    default:
                        MessageBox.Show("Bu hedef format için dönüşüm henüz desteklenmiyor.");
                        break;
                }
            }
        }

        private void HandleTxtConversion(string txtFilePath, string targetExtension)
        {
            string convertedText = ConvertWordToText(txtFilePath);

            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(txtFilePath);
            string newFileName = fileNameWithoutExtension + targetExtension;

            saveFileDialog1.FileName = newFileName;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string saveFilePath = saveFileDialog1.FileName;

                switch (targetExtension)
                {
                    case ".docx":
                        ConvertTxtToDocx(txtFilePath, saveFilePath);
                        MessageBox.Show("TXT başarıyla DOCX'e dönüştürüldü.");
                        break;
                    case ".pdf":
                        convertedText = ConvertDocxToPdf(txtFilePath, saveFilePath);
                        MessageBox.Show("TXT başarıyla PDF'ye dönüştürüldü.");
                        break;
                    case ".jpg":
                        ConvertTxtToJpg(txtFilePath, Path.GetDirectoryName(saveFilePath));
                        MessageBox.Show("TXT başarıyla JPG'e dönüştürüldü.");
                        break;
                    case ".png":
                        ConvertTxtToPng(txtFilePath, Path.GetDirectoryName(saveFilePath));
                        MessageBox.Show("TXT başarıyla PNG'e dönüştürüldü.");
                        break;
                    default:
                        MessageBox.Show("Bu hedef format için dönüşüm henüz desteklenmiyor.");
                        break;
                }
            }
        }

        private void HandleJpgConversion(string jpgFilePath, string targetExtension)
        {
            string convertedText = ConvertWordToText(jpgFilePath);

            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(jpgFilePath);
            string newFileName = fileNameWithoutExtension + targetExtension;

            saveFileDialog1.FileName = newFileName;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string saveFilePath = saveFileDialog1.FileName;

                switch (targetExtension)
                {
                    case ".png":
                        ConvertJpgToPng(jpgFilePath, Path.GetDirectoryName(saveFilePath)); // JPG'yi PNG'ye dönüştürme işlemi burada çağrılıyor.
                        MessageBox.Show("JPG başarıyla PNG'ye dönüştürüldü.");
                        break;
                    case ".ico":
                        ConvertJpgToIco(jpgFilePath, saveFilePath); // JPG'yi ICO'ya dönüştürme işlemi burada çağrılıyor.
                        MessageBox.Show("JPG başarıyla ICO'ya dönüştürüldü.");
                        break;
                    case ".pdf":
                        ConvertJpgToPdf(new string[] { jpgFilePath }, saveFilePath); // JPG'yi PDF'ye dönüştürme işlemi burada çağrılıyor.
                        MessageBox.Show("JPG başarıyla PDF'ye dönüştürüldü.");
                        break;
                    case ".docx":
                        convertedText = ConvertJpgToTxt(jpgFilePath); // JPG'yi TXT'ye dönüştürme işlemi burada çağrılıyor.
                        if (!string.IsNullOrEmpty(convertedText))
                        {
                            File.WriteAllText(saveFilePath, convertedText);
                            MessageBox.Show("JPG başarıyla metin (TXT) formatına dönüştürüldü ve kaydedildi.");
                        }
                        else
                        {
                            MessageBox.Show("JPG dosyası metin olarak dönüştürülemedi veya boş bir metin elde edildi.");
                        }
                        break;
                    case ".txt":
                        ConvertTxtToPng(jpgFilePath, Path.GetDirectoryName(saveFilePath));
                        MessageBox.Show("TXT başarıyla PNG'e dönüştürüldü.");
                        break;
                    default:
                        MessageBox.Show("Bu hedef format için dönüşüm henüz desteklenmiyor.");
                        break;
                }
            }
        }

        private void HandlePngConversion(string pngFilePath, string targetExtension)
        {
            string convertedText = ConvertWordToText(pngFilePath);

            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(pngFilePath);
            string newFileName = fileNameWithoutExtension + targetExtension;

            saveFileDialog1.FileName = newFileName;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string saveFilePath = saveFileDialog1.FileName;

                switch (targetExtension)
                {
                    case ".jpg":
                        ConvertPngToJpg(pngFilePath, Path.GetDirectoryName(saveFilePath)); // PNG'den JPG'ye dönüşüm işlemi burada çağrılıyor.
                        MessageBox.Show("PNG başarıyla JPG'ye dönüştürüldü.");
                        break;
                    case ".pdf":
                        ConvertPngToPdf(new string[] { pngFilePath }, saveFilePath); // PNG'yi PDF'ye dönüştürme işlemi burada çağrılıyor.
                        MessageBox.Show("PNG başarıyla PDF'ye dönüştürüldü.");
                        break;
                    case ".txt":
                        convertedText = ConvertJpgToTxt(pngFilePath); // JPG'yi TXT'ye dönüştürme işlemi burada çağrılıyor.

                        if (!string.IsNullOrEmpty(convertedText))
                        {
                            File.WriteAllText(saveFilePath, convertedText);
                            MessageBox.Show("JPG başarıyla metin (TXT) formatına dönüştürüldü ve kaydedildi.");
                        }
                        else
                        {
                            MessageBox.Show("JPG dosyası metin olarak dönüştürülemedi veya boş bir metin elde edildi.");
                        }
                        break;

                    case ".docx":
                        ConvertJpgToDocx(pngFilePath, saveFilePath); // JPGPNGyi DOCX'ye dönüştürme işlemi burada çağrılıyor.
                        MessageBox.Show("PNG başarıyla DOCX'e dönüştürüldü.");
                        break;
                    case ".ico":
                        try
                        {
                            ConvertPngToIco(pngFilePath, saveFilePath); // PNG dosyasını ICO'ya dönüştürme işlemi burada çağrılıyor.
                            MessageBox.Show("PNG başarıyla ICO'ya dönüştürüldü.");
                        }
                        catch (Exception ex)
                        {
                            HandleError("PNG'den ICO'ya dönüşüm sırasında bir hata oluştu: " + ex.Message);
                        }
                        break;
                    default:
                        MessageBox.Show("Bu hedef format için dönüşüm henüz desteklenmiyor.");
                        break;
                }
            }
        }

        private void HandleError(string errorMessage)
        {
            MessageBox.Show(errorMessage, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private string GetTesseractLanguageCode(string selectedLanguage)
        {
            switch (selectedLanguage)
            {
                case "Türkçe":
                    return "tur";
                case "English":
                    return "eng";
                case "Русский":
                    return "rus";
                default:
                    return "tur"; // Varsayılan olarak Turkçeyi kullan
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string selectedFilePath = openFileDialog1.FileName;
            string selectedExtension = Path.GetExtension(selectedFilePath);

            if (string.IsNullOrEmpty(selectedFilePath) || string.IsNullOrEmpty(selectedExtension))
            {
                MessageBox.Show("Lütfen bir dosya seçiniz!");
                return;
            }

            string targetExtension = comboBox3.SelectedItem?.ToString();

            if (targetExtension != null)
            {
                try
                {
                    switch (selectedExtension)
                    {
                        case ".jpg":
                        case ".png":
                            HandleImageConversion(selectedFilePath, selectedExtension, targetExtension);
                            break;
                        default:
                            MessageBox.Show("Bu uzantı için dönüşüm henüz desteklenmiyor.");
                            break;
                    }
                }
                catch (Exception ex)
                {
                    HandleError("Dönüşüm sırasında bir hata oluştu: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Lütfen bir hedef uzantı seçiniz!");
            }
        }

        private void HandleImageConversion(string imagePath, string selectedExtension, string targetExtension)
        {
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(imagePath);
            string newFileName = fileNameWithoutExtension + targetExtension;

            saveFileDialog1.FileName = newFileName;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string saveFilePath = saveFileDialog1.FileName;

                switch (targetExtension)
                {
                    case ".docx":
                        ConvertJpgToDocxLanguage(imagePath, saveFilePath, comboBox5.SelectedItem?.ToString());
                        MessageBox.Show("Metin başarıyla DOCX formatına dönüştürüldü ve kaydedildi.");
                        break;
                    case ".pdf":
                        if (selectedExtension == ".jpg")
                        {
                            ConvertJpgToPdf(new string[] { imagePath }, saveFilePath);
                        }
                        else if (selectedExtension == ".png")
                        {
                            ConvertPngToPdf(new string[] { imagePath }, saveFilePath);
                        }
                        MessageBox.Show("Dosya başarıyla PDF'ye dönüştürüldü.");
                        break;
                    case ".txt":
                        string convertedText = ConvertJpgToTxtLanguage(imagePath, comboBox5.SelectedItem?.ToString());
                        if (!string.IsNullOrEmpty(convertedText))
                        {
                            File.WriteAllText(saveFilePath, convertedText);
                            MessageBox.Show("Metin başarıyla TXT formatına dönüştürüldü ve kaydedildi.");
                        }
                        else
                        {
                            MessageBox.Show("Dosya metin olarak dönüştürülemedi veya boş bir metin elde edildi.");
                        }
                        break;
                    default:
                        MessageBox.Show("Bu hedef format için dönüşüm henüz desteklenmiyor.");
                        break;
                }
            }
        }

    }
}
