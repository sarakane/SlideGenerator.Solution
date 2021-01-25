using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Drawing;
using System.Net.Http;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.PowerPoint.Application;

namespace SlideGenerator
{
    public partial class CreateSlideForm : Form
    {
        private readonly List<string> selectedWords = new List<string>();
        public CreateSlideForm()
        {
            InitializeComponent();
        }

        private void CreateSlideButton_Click(object sender, System.EventArgs e)
        {
            Application pptApplication = new Application();
            Presentation pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);

            CustomLayout customLayout = pptPresentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutText];

            Slides slides = pptPresentation.Slides;
            _Slide slide = slides.AddSlide(1, customLayout);

            //Create Title from slideTitleTextBox
            TextRange slideTitle = slide.Shapes[1].TextFrame.TextRange;
            slideTitle.Text = slideTitleTextBox.Text;
            slideTitle.Font.Size = 32;
            slide.Shapes[1].Height = 60;

            //Create Text from slideTextRichTextBox
            TextRange slideText = slide.Shapes[2].TextFrame.TextRange;
            slideText.Text = slideTextRichTextBox.Text;
            slideText.Font.Size = 16;
            slide.Shapes[2].Width = 310;
            slide.Shapes[2].Top = 115;
        }

        private void BoldTextButton_Click(object sender, System.EventArgs e)
        {
            if (slideTextRichTextBox.SelectionFont.Bold)
            {
                slideTextRichTextBox.SelectionFont = new System.Drawing.Font("Microsoft Sans Serif", (float)8.25, FontStyle.Regular, GraphicsUnit.Point);
                selectedWords.Remove(slideTextRichTextBox.SelectedText.Trim(' ', ',', '.', '!', '?'));
            }
            else
            {
                slideTextRichTextBox.SelectionFont = new System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point);
                selectedWords.Add(slideTextRichTextBox.SelectedText.Trim(' ', ',', '.', '!', '?'));
            }
        }

        private async void GetImagesButton_ClickAsync(object sender, System.EventArgs e)
        {
            // Images are fetched using the Bing Image Search API 
            string _subscriptionKey = Credentials.ApiKey; // API key is stored in a class that git is set to ignore
            string baseUri = "https://api.bing.microsoft.com/v7.0/images/search";
            string mkt_parameter = "&mkt=en-us";
            string count_parameter = "&count=10";
            string query_parameter = "?q=";

            string titleSearch = slideTitleTextBox.Text;

            query_parameter += titleSearch + " ";
            foreach (string word in selectedWords)
            {
                query_parameter += "OR " + word;
            }

            string requestUri = baseUri + query_parameter + mkt_parameter + count_parameter;

            var client = new HttpClient();
            client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", _subscriptionKey);
            string myData = await client.GetStringAsync(requestUri);
            dynamic convert = JsonConvert.DeserializeObject(myData);
            dynamic images = convert["value"];
            pictureBox1.ImageLocation = images[0]["contentUrl"].ToString();
            pictureBox2.ImageLocation = images[1]["contentUrl"].ToString();
            pictureBox3.ImageLocation = images[2]["contentUrl"].ToString();
            pictureBox4.ImageLocation = images[3]["contentUrl"].ToString();
            pictureBox5.ImageLocation = images[4]["contentUrl"].ToString();
        }

    }
}
