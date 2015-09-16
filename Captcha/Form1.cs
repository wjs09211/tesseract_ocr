using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using mshtml;

namespace Captcha
{
    public partial class Form1 : Form
    {
        Image img;
        Bitmap bimg;
        string Security_Code;
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
        }
        private void button1_Click(object sender, EventArgs e)
        {
            webBrowser1.Navigate("https://isdna1.yzu.edu.tw/CnStdSel/SelRandomImage.aspx");
        }
        //讀入驗證碼
        private Image Getimage()
        {
            HTMLDocument doc = webBrowser1.Document.DomDocument as HTMLDocument;
            HTMLBody body = doc.body as HTMLBody;
            IHTMLControlRange range = body.createControlRange();
            //取得網頁中第[0]個圖片(因為我要取得的圖片的沒有ID)
            //也可使用GetElementById("ID名稱")取得圖片
            IHTMLControlElement imgElement =
                webBrowser1.Document.GetElementsByTagName("img")[0].DomElement as IHTMLControlElement;
            //複製圖片至剪貼簿
            range.add(imgElement);
            range.execCommand("copy", false, Type.Missing);
            //取得圖片
            img = Clipboard.GetImage();
            Clipboard.Clear();
            return img;
        }
        public string execution()
        {
            Image lookimg;
            //轉灰階
            convert2GrayScale();
            lookimg = Image.FromHbitmap(bimg.GetHbitmap());
            lookimg = new Bitmap(lookimg, lookimg.Width * 2, lookimg.Height * 2);
            pictureBox2.Image = lookimg;
            //去雜線
            RemoteNoiseLineByPixels();
            lookimg = Image.FromHbitmap(bimg.GetHbitmap());
            lookimg = new Bitmap(lookimg, lookimg.Width * 2, lookimg.Height * 2);
            pictureBox3.Image = lookimg;
            //去外框
            ClearPictureBorder(2);
            //去雜點
            RemoteNoisePointByPixels();
            lookimg = Image.FromHbitmap(bimg.GetHbitmap());
            lookimg = new Bitmap(lookimg, lookimg.Width * 2, lookimg.Height * 2);
            pictureBox4.Image = lookimg;
            //補點
            AddNoisePointByPixels();
            lookimg = Image.FromHbitmap(bimg.GetHbitmap());
            lookimg = new Bitmap(lookimg, lookimg.Width * 2, lookimg.Height * 2);
            pictureBox5.Image = lookimg;
            img = Image.FromHbitmap(bimg.GetHbitmap());
            parseCaptchaStr(img);
            return Security_Code;
        }
        private void convert2GrayScale()
        {
            for (int i = 0; i < bimg.Width; i++)
            {
                for (int j = 0; j < bimg.Height; j++)
                {
                    Color pixelColor = bimg.GetPixel(i, j);
                    byte r = pixelColor.R;
                    byte g = pixelColor.G;
                    byte b = pixelColor.B;

                    byte gray = (byte)(0.299 * (float)r + 0.587 * (float)g + 0.114 * (float)b);
                    r = g = b = gray;
                    pixelColor = Color.FromArgb(r, g, b);

                    bimg.SetPixel(i, j, pixelColor);
                }
            }
        }
        //去邊框
        private void ClearPictureBorder(int pBorderWidth)
        {
            for (int i = 0; i < bimg.Height; i++)
            {
                for (int j = 0; j < bimg.Width; j++)
                {
                    if (i < pBorderWidth || j < pBorderWidth || j > bimg.Width - 1 - pBorderWidth || i > bimg.Height - 1 - pBorderWidth)
                        bimg.SetPixel(j, i, Color.FromArgb(255, 255, 255));
                }
            }
        }
        //去雜點
        private class NoisePoint
        {
            public int X { get; set; }
            public int Y { get; set; }
        }
        private void RemoteNoisePointByPixels()
        {
            List<NoisePoint> points = new List<NoisePoint>();

            for (int k = 0; k < 5; k++)
            {
                for (int i = 0; i < bimg.Height; i++)
                    for (int j = 0; j < bimg.Width; j++)
                    {
                        int flag = 0;
                        int garyVal = 255;
                        // 檢查上相鄰像素
                        if (i - 1 > 0 && bimg.GetPixel(j, i - 1).R != garyVal) flag++;
                        if (i + 1 < bimg.Height && bimg.GetPixel(j, i + 1).R != garyVal) flag++;
                        if (j - 1 > 0 && bimg.GetPixel(j - 1, i).R != garyVal) flag++;
                        if (j + 1 < bimg.Width && bimg.GetPixel(j + 1, i).R != garyVal) flag++;
                        if (i - 1 > 0 && j - 1 > 0 && bimg.GetPixel(j - 1, i - 1).R != garyVal) flag++;
                        if (i + 1 < bimg.Height && j - 1 > 0 && bimg.GetPixel(j - 1, i + 1).R != garyVal) flag++;
                        if (i - 1 > 0 && j + 1 < bimg.Width && bimg.GetPixel(j + 1, i - 1).R != garyVal) flag++;
                        if (i + 1 < bimg.Height && j + 1 < bimg.Width && bimg.GetPixel(j + 1, i + 1).R != garyVal) flag++;

                        if (flag < 3)
                            points.Add(new NoisePoint() { X = j, Y = i });
                    }
                foreach (NoisePoint point in points)
                    bimg.SetPixel(point.X, point.Y, Color.FromArgb(255, 255, 255));

            }
        }
        //去噪音線
        private void RemoteNoiseLineByPixels()
        {
            for (int i = 0; i < bimg.Height; i++)
                for (int j = 0; j < bimg.Width; j++)
                {
                    int grayValue = bimg.GetPixel(j, i).R;
                    if (grayValue <= 255 && grayValue >= 160)
                        bimg.SetPixel(j, i, Color.FromArgb(255, 255, 255));
                }
        }
        //補點
        private void AddNoisePointByPixels()
        {
            List<NoisePoint> points = new List<NoisePoint>();

            for (int k = 0; k < 1; k++)
            {
                for (int i = 0; i < bimg.Height; i++)
                    for (int j = 0; j < bimg.Width; j++)
                    {
                        int flag = 0;
                        int garyVal = 255;
                        // 檢查上相鄰像素
                        if (i - 1 > 0 && bimg.GetPixel(j, i - 1).R != garyVal) flag++;
                        if (i + 1 < bimg.Height && bimg.GetPixel(j, i + 1).R != garyVal) flag++;
                        if (j - 1 > 0 && bimg.GetPixel(j - 1, i).R != garyVal) flag++;
                        if (j + 1 < bimg.Width && bimg.GetPixel(j + 1, i).R != garyVal) flag++;
                        if (i - 1 > 0 && j - 1 > 0 && bimg.GetPixel(j - 1, i - 1).R != garyVal) flag++;
                        if (i + 1 < bimg.Height && j - 1 > 0 && bimg.GetPixel(j - 1, i + 1).R != garyVal) flag++;
                        if (i - 1 > 0 && j + 1 < bimg.Width && bimg.GetPixel(j + 1, i - 1).R != garyVal) flag++;
                        if (i + 1 < bimg.Height && j + 1 < bimg.Width && bimg.GetPixel(j + 1, i + 1).R != garyVal) flag++;

                        if (flag >= 7)
                            points.Add(new NoisePoint() { X = j, Y = i });
                    }
                foreach (NoisePoint point in points)
                    bimg.SetPixel(point.X, point.Y, Color.FromArgb(0, 0, 0));
            }
        }


        private void parseCaptchaStr(Image image)
        {
            Ocr ocr = new Ocr();
            Bitmap BmpSource = new Bitmap(image);
            ocr.DoOCRMultiThred(BmpSource, "eng");
            Security_Code = ocr.GetSecurity_Code();
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            img = Getimage();
            pictureBox1.Image = img;
            bimg = new Bitmap(img);
            textBox1.Text = execution();
        }
    }
}
