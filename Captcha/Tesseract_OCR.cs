using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Drawing;
using System.Windows.Forms;
namespace Captcha
{
    public class Ocr
    {
        string Security_Code;
        public void DumpResult(List<tessnet2.Word> result)
        {
            var sb = new StringBuilder();
            result.ForEach(c =>
            {
                sb.Append(c.Text);
            });
            Security_Code = sb.ToString();
        }
        public string GetSecurity_Code()
        {
            return Security_Code;
        }
        public List<tessnet2.Word> DoOCRNormal(Bitmap image, string lang)
        {
            tessnet2.Tesseract ocr = new tessnet2.Tesseract();
            ocr.Init(null, lang, false);
            List<tessnet2.Word> result = ocr.DoOCR(image, Rectangle.Empty);
            DumpResult(result);
            return result;
        }

        ManualResetEvent m_event;

        public void DoOCRMultiThred(Bitmap image, string lang)
        {
            tessnet2.Tesseract ocr = new tessnet2.Tesseract();
            ocr.SetVariable("tessedit_char_whitelist", "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789");
            var path = string.Concat(Application.StartupPath, @"\tessdata");
            ocr.Init(path, lang, false);
            // If the OcrDone delegate is not null then this'll be the multithreaded version
            ocr.OcrDone = new tessnet2.Tesseract.OcrDoneHandler(Finished);
            // For event to work, must use the multithreaded version
            ocr.ProgressEvent += new tessnet2.Tesseract.ProgressHandler(ocr_ProgressEvent);
            m_event = new ManualResetEvent(false);
            ocr.DoOCR(image, Rectangle.Empty);
            // Wait here it's finished
            m_event.WaitOne();
        }

        public void Finished(List<tessnet2.Word> result)
        {
            DumpResult(result);
            m_event.Set();
        }

        void ocr_ProgressEvent(int percent)
        {
            Console.WriteLine("{0}% progression", percent);
        }
    }
}
