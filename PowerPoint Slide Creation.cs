using System;
using System.Drawing;
using System.Net;
using System.Net.Http;
using System.Configuration;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using Google.Apis.CustomSearchAPI.v1;
using PPT = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.IO;

namespace WindowsFormsApplication1
{

    public partial class PowerPointSlide : Form
    {
        // I'd make a call to a secured database or something but considering the scope of this that's overkill
        string apiKey = "AIzaSyDVB3vjtC721EaDw_KFkg0utQq7HAF-quU";
        string customSearch = "afa084a1c2d0eeb3e";

        WebClient downloader = new WebClient();
        HttpClient client = new HttpClient();

        Button[] imageGrid;
        GoogleImageItems images;
        List<int> selected;

        public PowerPointSlide()
        {
            InitializeComponent();
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {
            
        }

        private void PowerPointSlide_Load(object sender, EventArgs e)
        {
            // Initialize shared lists
            imageGrid = new Button[] { imagebutton1, imagebutton2, imagebutton3, imagebutton4, imagebutton5, imagebutton6, imagebutton7, imagebutton8, imagebutton9, imagebutton10 };
            selected = new List<int>();

        }

        #region Toolbar Buttons
        private void BoldButton_Click(object sender, EventArgs e)
        {
            Font toggled, old;
            old = richTextBox1.SelectionFont;
            if(old.Bold)
            {
                toggled = new Font(old, old.Style & ~FontStyle.Bold);
            }
            else
            {
                toggled = new Font(old, old.Style | FontStyle.Bold);
            }

            richTextBox1.SelectionFont = toggled;
            richTextBox1.Focus();
        }

        private void ItalicsButton_Click(object sender, EventArgs e)
        {
            Font toggled, old;
            old = richTextBox1.SelectionFont;
            if (old.Bold)
            {
                toggled = new Font(old, old.Style & ~FontStyle.Italic);
            }
            else
            {
                toggled = new Font(old, old.Style | FontStyle.Italic);
            }

            richTextBox1.SelectionFont = toggled;
            richTextBox1.Focus();
        }
        #endregion

        // Function is called upon button press and searches for 10 images related to the title and bolded text in the description
        private async void SearchForImages(object sender, EventArgs e)
        {
            try
            {
                selected = new List<int>();
                foreach (var image in imageGrid)
                {
                    image.BackColor = SystemColors.MenuBar;
                }

                string query = textBox1.Text;

                // Logic for finding bolded words
                List<int> spaces = new List<int>() { 0 };
                char[] space = new char[] { ' ', ',', '.'};
                int i;

                //Loop through and find all the spaces in text
                while ((i = richTextBox1.Find(space, spaces.Last() + 1)) != -1) { spaces.Add(i); }
                spaces[0] = -1;
                spaces.Add(richTextBox1.TextLength);

                // Check each area between spaces and see if it's bolded. If so add that to the query terms
                for (i = 0; i < spaces.Count - 1; i++)
                {
                    // ignore bolded spaces :)
                    if (spaces[i] == spaces[i + 1] - 1) { continue; }

                    richTextBox1.Select(spaces[i] + 1, spaces[i + 1] - spaces[i] - 1);
                    if (richTextBox1.SelectionFont.Bold) { query += " " + richTextBox1.SelectedText; }
                }

                // Make query uri friendly
                query = query.Replace(' ', '%');

                //ignore search if nothing is being searched
                if (query == "") { return; }

                // Search online for keywords
                images = await Search(query);

                // if no results are found show a blank icon
                if (images.items == null)
                {
                    for (int c = 0; c < imageGrid.Length; c++)
                    {
                        using (var tempBitmap = new Bitmap("blank_image.jpg"))
                        {
                            imageGrid[c].BackgroundImage = new Bitmap(tempBitmap);
                        }
                    }
                }
                DownloadIcons(images);

                // Display the thumbnail images
                for (int c = 0; c < imageGrid.Length; c++)
                {
                    using (var tempBitmap = new Bitmap(string.Format("thumb{0}.jpg", c)))
                    {
                        imageGrid[c].BackgroundImage = new Bitmap(tempBitmap);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        #region helperfunctions
        private async Task<GoogleImageItems> Search(string query)
        {
            string url = string.Format("https://customsearch.googleapis.com/customsearch/v1?cx={0}&exactTerms={1}&searchType=image&key={2}", customSearch, query, apiKey);
            var searchRequest = await client.GetStringAsync(url);

            var json = JsonConvert.DeserializeObject<GoogleImageItems>(searchRequest);

            return json;
        }

        private void DownloadIcons(GoogleImageItems google_images)
        {
            for (int c = 0; c < imageGrid.Length; c++)
            {
                downloader.DownloadFile(
                    google_images.items[c].image.thumbnailLink, 
                    string.Format(@"thumb{0}.jpg", c));
            }

            return;
        }

        private void SelectImage(object sender, EventArgs e)
        {
            // Converst generic types to specific types
            Button realSender = (Button)sender;
            int id = int.Parse((string)realSender.Tag) - 1;

            if (selected.Contains(id))
            {
                selected.Remove(id);
                realSender.BackColor = SystemColors.MenuBar;
            }
            else
            {
                selected.Add(id);
                realSender.BackColor = SystemColors.MenuHighlight;
            }
        }
        #endregion

        private void CreateSlide(object sender, EventArgs e)
        {
            //initialize PowerPoint
            PPT.Application pptApplication = new PPT.Application();
            PPT.Presentation pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);
            try
            {
                // My version of PowerPoint doesn't seem to match up with the version expected here so I hardcoded the number that matches my content with caption layout
                // I think The commented out line of code below works in Office 365
                //var customLayout = pptPresentation.SlideMaster.CustomLayouts[PPT.PpSlideLayout.ppLayoutContentWithCaption];
                var customLayout = pptPresentation.SlideMaster.CustomLayouts[8];
                var slide = pptPresentation.Slides.AddSlide(1, customLayout);

                // Add Title text
                var title = slide.Shapes[1].TextFrame.TextRange;
                title.Text = this.textBox1.Text;
                title.Font.Name = "Arial";
                title.Font.Size = 46;

                // Add Content text
                var content = slide.Shapes[3].TextFrame.TextRange;
                content.Text = this.richTextBox1.Text;

                //Get parameters of where images should be placed based on third box then remove that box
                PPT.Shape imageBox = slide.Shapes[2];
                float left = imageBox.Left;
                float top = imageBox.Top;
                float width = imageBox.Width / 2;
                float height = imageBox.Height / 2;
                int bit = 0;
                slide.Shapes[2].Delete();

                string fullPath = Directory.GetCurrentDirectory();

                foreach (int i in selected)
                {
                    // Only allow 4 pictures to be added since there isn't really space for more
                    if (bit == 4) { break; }
                    string fileName = string.Format(@"{0}\\full{1}.jpg", fullPath, i);

                    // download the image and load it into the presentation. If the image can't be downloaded from the site skip it gracefully
                    try { downloader.DownloadFile(images.items[i].link, fileName); }
                    catch { continue; }
                    slide.Shapes.AddPicture(
                        fileName, 
                        MsoTriState.msoFalse, 
                        MsoTriState.msoTrue, 
                        // Isn't bit shifting fun? The below code aligns the 4 pictures to be laid out in a square
                        left+((bit&1)*width), 
                        top+((bit>>1)*height), 
                        width, 
                        height);

                    bit++;
                }

                //Save the file to the user's documents folder
                var documentsPath = string.Format(@"{0}\\trialPowerPoint.pptx", Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                pptPresentation.SaveAs(documentsPath, PPT.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);

                // Make sure we always close PowerPoint after finishing
                pptPresentation.Close();
                pptApplication.Quit();

                resultLabel.Text = "The powerpoint slide has been successfully created and saved to your Documents folder as 'trialPowerPoint'";
            }
            // Catch any errors, let the user know it didn't go through, and make sure the powerpoint is cleaned up
            catch (Exception ex)
            {
                // Make sure we always close PowerPoint after finishing
                pptPresentation.Close();
                pptApplication.Quit();
                resultLabel.Text = "Sorry, we ran into an issue and were unable to create your PowerPoint Slide.";
                Console.WriteLine(ex);
            }
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }
    }
}
