using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.PowerPoint.Application;

namespace SlideGenerator
{
    public partial class CreateSlideForm : Form
    {
        public CreateSlideForm()
        {
            InitializeComponent();
        }

        private void CreateSlideButton_Click(object sender, System.EventArgs e)
        {
            Application pptApplication = new Application();
            Presentation pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);

            CustomLayout customLayout = pptPresentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutText];

            var slides = pptPresentation.Slides;
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
    }
}
