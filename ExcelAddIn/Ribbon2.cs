using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Diagnostics;
using System.Speech;
using Microsoft.Office.Interop.Excel;
using System.Speech.Synthesis;
using System.Drawing;
using System.Windows.Forms;
using System.Threading.Tasks;

namespace ExcelAddIn
{
    public partial class Ribbon2
    {
        private void Ribbon2_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btn_help_Click(object sender, RibbonControlEventArgs e)
          //câu 1
          //Dùng hàm Process.Start cùng thư viện System.Diagnostics để chạy đến web của SoICT
        { //redirected to SoICT's website
            Process.Start(@"https://soict.hust.edu.vn");
        }
        SpeechSynthesizer speechSynthesizerObj;
        private void btn_TextToSpeech_Click(object sender, RibbonControlEventArgs e)
        {
            //câu 2
            Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            //Chọn cell
            Range currentRange = (Range)Globals.ThisAddIn.Application.Selection as
            Microsoft.Office.Interop.Excel.Range;
            if (currentRange == null) return; //tránh range rỗng
            foreach (string mycell in currentRange.Cells)
            {
                //ignore null cells
                if (mycell != null && ((dynamic)(mycell)).Value != null)
                //nếu giá trị của cells khác rỗng
                {
                    speechSynthesizerObj = new SpeechSynthesizer();
                    //tạo object SpeechSynthesizer
                    speechSynthesizerObj.SetOutputToDefaultAudioDevice();
                    //đặt output âm thanh đầu ra default
                    speechSynthesizerObj.Speak(mycell);
                    //nói giá trị của Value
                }
            }
        }

        private void dropDown1_SelectionChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void btn_RGB_Click(string color, string saturation)
        { 
            //câu 3(phần comment)
            //get selected cells
            Range currentRange = (Range)Globals.ThisAddIn.Application.Selection as
                Microsoft.Office.Interop.Excel.Range;
            if (currentRange == null) return;

            //Đọc từng cell và tô màu
            //foreach (var mycell in currentRange.Cells)
            //{
            //    //Nếu giá trị của cell khác rỗng thì
            //    if (mycell != null && ((dynamic)(mycell)).Value != null)
            //    {
            //        (dynamic)(mycell)).Value = int saturationInt;
            //      //đặt giá trị value là một biến tên saturationInt, giá trị int
            //        currentRange.Interior.Color = Color.FromArgb(saturationInt, saturationInt, saturationInt);
            //      //chuyển màu của biến thành màu đa mức xám
            //        if (saturationInt > 128) // nếu biến lớn hơn 128 ta sẽ chuyển font color thành đen
            //        {
            //            currentRange.Font.Color = Color.Black;
            //        }
            //        else //ngược lại ta sẽ chuyển font color thành màu trắng
            //        {
            //            currentRange.Font.Color = Color.White;
            //        }
            //    }
            //}
            //hết câu 3
            // nhược điểm: chưa đọc được ví dụ nếu số nhập vào là lỗi hoặc lớn hơn 255.
            int saturationInt; //saturation value in integer 
            bool isNumeric = int.TryParse(  saturation, out saturationInt); //boolean to check if saturation is a valid number
            if (isNumeric) //check if the saturation is a number
            {
                if (saturationInt >= 0 && saturationInt <= 255) //check if the saturation value is in range 0 to 255
                {
                    switch (color) // color to display with saturation 
                    {
                        case "AppointmentColor1": currentRange.Interior.Color = Color.FromArgb(saturationInt, 0, 0); break; //red
                        case "AppointmentColor2": currentRange.Interior.Color = Color.FromArgb(0, 0, saturationInt); break; //green
                        case "AppointmentColor3": currentRange.Interior.Color = Color.FromArgb(0, saturationInt, 0); break; //blue
                        case "AppointmentColor4": currentRange.Interior.Color = Color.FromArgb(saturationInt, saturationInt, saturationInt); break; //gray
                        default: break;
                    }
                    if (saturationInt > 128) // change font color to black or white based on background's saturation
                    {
                        currentRange.Font.Color = Color.Black;
                    }
                    else
                    {
                        currentRange.Font.Color = Color.White;
                    }
                    currentRange.Value = saturationInt; // cell value = saturation value
                }
                else
                {
                    //saturation is not in range 0 to 255
                    MessageBox.Show("Please input a number between 0 and 255 at saturation box", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                //saturation is not a valid number
                MessageBox.Show("Invalid type, please input a number at saturation box", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
    }
}
