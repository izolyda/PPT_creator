using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using System.Windows.Markup;
using System.Xml.Linq;
using System.Xml;
using System.Text.RegularExpressions;

using System.Threading.Tasks;

using Google.Apis.Discovery;
using Google.Apis.Services;
using System.Net;
using Newtonsoft.Json.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Diagnostics;

using Microsoft.Office.Interop;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using System.Xaml;
using System.Drawing.Imaging;


namespace PPT_creator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        string APIkey = "AIzaSyD1D-imowqet24jydMF3kX4f17vBIWP0w8";
        string cx="a5a010075cde35c18";

        List<String> keywords = new List<String>();
        SlidesCollectionHelper slidesCollectionHelper = new SlidesCollectionHelper();
        PicsCollectionHelper picsCollectionHelper = new PicsCollectionHelper();

        static Application pptApplication = new Application();
        Presentation pptPresentation;
        Microsoft.Office.Interop.PowerPoint.CustomLayout customLayout;

        string picDir = "C:\\temp\\Images";
        

        byte[] imageBytes = null;

        Microsoft.Office.Interop.PowerPoint.Slides slides = null;



        public MainWindow()
        {
            InitializeComponent();
           
            mainRTB.AllowDrop = true;

            Init();
        }

        private void Init()
        {
            if(pptPresentation!=null)
            {
                pptPresentation.Close();
            }

            pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);
            customLayout = pptPresentation.SlideMaster.CustomLayouts[Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutClipartAndText];

            string fileName = "presentation.pptx";
            string pptPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);

            try
            {
                pptPresentation.SaveAs(pptPath, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
            } catch(Exception ex)
            {
                MessageBox.Show("Couldn't create presentation at this time. Please close the file if open in other applications, exit and try again.");
            }
           

            System.IO.Directory.CreateDirectory(@picDir);

        }


        private void getTitleKeywords()
        {
            string title = titleArea.Text;
            string[] words = title.Split(' ');

            foreach (var el in words)
            {
                if(!keywords.Contains(el))
                    keywords.Add(el);
            }
        }
       

        private void getKeywords()
        {
            List<String> tmp = new List<String>();

            System.Windows.Documents.TextRange allText = new System.Windows.Documents.TextRange(mainRTB.Document.ContentStart, mainRTB.Document.ContentEnd);

            MemoryStream memstream = new MemoryStream();
            allText.Save(memstream, DataFormats.Xaml);
            if (memstream != null)
            {
                memstream.Close();
            }

            string rawxaml = Encoding.ASCII.GetString(memstream.ToArray());


            string regex = "(FontWeight=\\\"Bold\\\").*?>(\\w.*?)(<\\/Run>)";
            MatchCollection coll = Regex.Matches(rawxaml, regex);

            String result = "";

            if (coll.Count > 0)
            {
                for (int i = 0; i < coll.Count; i++)
                {
                    result = coll[i].Groups[2].Value;
                    tmp.Add(result);
                }

            }

            foreach (var el in tmp)
            {
                if (!keywords.Contains(el))
                {
                    keywords.Add(el);
                }


            }

            foreach (var el in keywords)
                Console.WriteLine(el);
        }


       
        private async void imageSearch(object sender, RoutedEventArgs e)
        {
            keywords.Clear();
            imagesStackPanel.Children.Clear();

            getTitleKeywords();
            getKeywords();

            string responseJSON = "";

            List<string> urls = new List<string>();

            IEnumerable<string> queryResults;

            try { 

            //Using a single client to make the calls.
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Accept.Clear();
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                  
                    //create all tasks
                    Task<string>[] alltasks = new Task<string>[keywords.Count];

                    for(int i=0; i<keywords.Count; i++)
                    {
                        alltasks[i] = GetAsync(keywords[i], 5);
                    }


                    // Await the completion of all the running tasks. 
                    var responses = await Task.WhenAll(alltasks); 

                    queryResults = responses.Where(r => r != null); //filter out any null values

                    foreach(var item in queryResults)
                    {
                        JObject jsonObj = JObject.Parse(item);
                        var imgUrls =
                               from lnk in jsonObj["items"]
                               select (string)lnk["link"];

                        int i = 1;
                        foreach (var l in imgUrls)
                        {
                            urls.Add(l);
                            Console.WriteLine(l);
                            appendImage(l, i);
                            i++;
                        }

                    }
                }
            }
            catch (AggregateException ex)
            {
                Console.WriteLine(ex);
            }
          
  
        }

       
        private async Task<string> GetAsync(string keyword, int limit)
        {
            string endpointURL = "https://www.googleapis.com/customsearch/v1";
            string searchType = "image";
            string filetype = "jpg";

            string result = "";

            var URL = "https://customsearch.googleapis.com/customsearch/v1?cx=" +
               cx
               + "&exactTerms="+ 
               keyword 
               +"&num=" + 
               limit +
               "&searchType=image" +
               "&fileType=" +
               filetype
               + "&imgSize=MEDIUM" +
               "&imgType=photo" +
               "&key=" +
               APIkey;

            HttpWebRequest myRequest = (HttpWebRequest)WebRequest.Create(URL);
            try
            {
                using (WebResponse myResponse = await myRequest.GetResponseAsync())
                {
                    using (StreamReader sr = new StreamReader(myResponse.GetResponseStream(), System.Text.Encoding.UTF8))
                    {
                        result = sr.ReadToEnd();
                    }
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex);
            }
   
            //Console.WriteLine(result);
            return result;
        }

        private void appendImage(string url, int i)
        {
            if (url != null)
            {            
                    StackPanel sp = new StackPanel();
                    sp.Margin = new Thickness(5);
                    sp.Name = "sp"+i;

                    Image img = new Image();
                    img.Width = 96;
                    img.Stretch = Stretch.Uniform;
                    img.Source = new BitmapImage(new Uri(url));
             
                    sp.Children.Add(img);

                    imagesStackPanel.Children.Add(sp);

                img.MouseMove += (s, e) => Img_MouseMove(s, e, url, imageBytes);
                
            }

        }

       

        private void Img_MouseMove(object sender, MouseEventArgs e, string url, byte[] imageBytes)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                //create an image byte array
                BitmapImage image = new BitmapImage(new Uri(url));
                var webClient = new WebClient();
                try
                {
                    imageBytes = webClient.DownloadData(url);
                    DataObject data = new DataObject();

                    data.SetData(imageBytes);

                    DragDrop.DoDragDrop(this, data, DragDropEffects.Copy | DragDropEffects.Move);

                }
                catch (Exception ex)
                {
                    Console.WriteLine("Forbidden.   reson:" + ex);
                }
                                
                Debug.WriteLine("Created.");
            }
        }

       
        private void DropEventHandler(object sender, DragEventArgs e)
        {

            if (Mouse.LeftButton == MouseButtonState.Released)
            {
                if (e.Data.GetDataPresent(typeof(Byte[])))
                {

                    byte[] imageBytes = e.Data.GetData(typeof(Byte[])) as Byte[];

                    
                    System.Drawing.Bitmap bmp;
                    using (var ms = new MemoryStream(imageBytes))
                    {
                        bmp = new System.Drawing.Bitmap(ms);
                    }

                    Image imgControl = new Image();
                    InlineUIContainer container = new InlineUIContainer(imgControl);
                    Paragraph paragraph = new Paragraph(container);


                    BitmapSource bs = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                      bmp.GetHbitmap(),
                      IntPtr.Zero,
                      System.Windows.Int32Rect.Empty,
                      BitmapSizeOptions.FromWidthAndHeight(96, 96));
                    ImageBrush ib = new ImageBrush(bs);
                    paragraph.Background = ib;

                    imgControl.Source = bs;

                    mainRTB.Document.Blocks.Add(paragraph);

                    Debug.WriteLine("Stop here.");

                   
                    if (e.KeyStates.HasFlag(DragDropKeyStates.ControlKey))
                    {
                        e.Effects = DragDropEffects.Copy;
                    }
                    else
                    {
                        e.Effects = DragDropEffects.Move;
                    }

                    e.Handled = true;

                }
               
                
            }     

        }

        private void mainRTB_Drop(object sender, DragEventArgs e)
        {
            if (Mouse.LeftButton == MouseButtonState.Released)
            {
                if (e.Data.GetDataPresent(typeof(Byte[])))
                {

                    byte[] imageBytes = e.Data.GetData(typeof(Byte[])) as Byte[];

                    System.Drawing.Bitmap bmp;
                    using (var ms = new MemoryStream(imageBytes))
                    {
                        bmp = new System.Drawing.Bitmap(ms);
                    }

                    Image imgControl = new Image();
                    
                    InlineUIContainer container = new InlineUIContainer(imgControl);
                    Paragraph paragraph = new Paragraph(container);


                    BitmapSource bs = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                      bmp.GetHbitmap(),
                      IntPtr.Zero,
                      System.Windows.Int32Rect.Empty,
                      BitmapSizeOptions.FromWidthAndHeight(bmp.Width/2, bmp.Height/2));

                    imgControl.Width = bmp.Width / 2;
                    imgControl.Height = bmp.Height / 2;
                    imgControl.Source = bs;

                    mainRTB.Document.Blocks.Add(paragraph);

                    //SAVE PIC TO COLLECTION AND DISK

                    PictureHelper pic = new PictureHelper();
                    picsCollectionHelper.addPic(pic);

                    string fileName = "image"+picsCollectionHelper.picsList.ElementAt(picsCollectionHelper.Count()-1).getPictureId() + ".jpg";

                    picsCollectionHelper.picsList.ElementAt(picsCollectionHelper.Count() - 1).setPictureUri(picDir+"\\"+fileName);


                    using (System.Drawing.Bitmap copyBmp = new System.Drawing.Bitmap(bmp.Width, bmp.Height))
                    {
                        using (System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(copyBmp))
                        {
                            g.Clear(System.Drawing.Color.White);
                            g.DrawImage(bmp, 0, 0, bmp.Width, bmp.Height);
                        }
                       
                        copyBmp.Save(picDir + "\\" + fileName, ImageFormat.Jpeg);
                        copyBmp.Dispose();
                    }

                    

                    bmp.Dispose();
                    

                    Debug.WriteLine("Stop here.");

                    
                    if (e.KeyStates.HasFlag(DragDropKeyStates.ControlKey))
                    {
                        e.Effects = DragDropEffects.Copy;
                    }
                    else
                    {
                        e.Effects = DragDropEffects.Move;
                    }

                    e.Handled = true;

                }


            }
        }

        private void mainRTB_PreviewDragOver(object sender, DragEventArgs e)
        {
            e.Handled = true;            
        }

        private void nextSlide(object sender, RoutedEventArgs e)
        {

            MemoryStream memstream = saveToMemStream();


            System.Windows.Documents.TextRange allText = new System.Windows.Documents.TextRange(mainRTB.Document.ContentStart, mainRTB.Document.ContentEnd);
             string filePath = @"thexaml.xaml";
             FileStream fs = new FileStream(filePath, FileMode.Create, FileAccess.Write);
             allText.Save(fs, DataFormats.Xaml);
             fs.Close();

            SlideHelper slideHelper = new SlideHelper(1);
            slidesCollectionHelper.addSlide(slideHelper);

            saveSlide(memstream);

            mainRTB.Document.Blocks.Clear();
            imagesStackPanel.Children.Clear();
            titleArea.Text = "placeholder";

            Console.WriteLine("stop.");
        }

        private MemoryStream saveToMemStream()
        {
            System.Windows.Documents.TextRange allText = new System.Windows.Documents.TextRange(mainRTB.Document.ContentStart, mainRTB.Document.ContentEnd);

            MemoryStream memstream = new MemoryStream();
            allText.Save(memstream, DataFormats.Xaml);

            return memstream;
        }

        private void SaveAll(object sender, RoutedEventArgs e)
        {
            pptPresentation.Close();
            pptApplication.Quit();

           
            MessageBox.Show("Successfully saved.");
            System.Windows.Application.Current.Shutdown();
            
                
            
        }


        private void saveSlide(MemoryStream ms)
        {
            Microsoft.Office.Interop.PowerPoint._Slide slide;
            Microsoft.Office.Interop.PowerPoint.TextRange objText;

           
            int id = slidesCollectionHelper.slides[0].getId();
            slidesCollectionHelper.slides.RemoveAt(0);

            try
            {
                slides = pptPresentation.Slides; 
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Data.Values.ToString());
            }

            slide = slides.AddSlide(id, customLayout); //goes to the collection of PPT slides

            slide.Layout = PpSlideLayout.ppLayoutText;


            // Add title
            objText = slide.Shapes[1].TextFrame.TextRange;
            objText.Text = titleArea.Text;
            objText.Font.Name = "Arial";
            objText.Font.Size = 32;


            string xaml = Encoding.ASCII.GetString(ms.ToArray());

            xaml = System.Text.RegularExpressions.Regex.Unescape(xaml);
            Console.WriteLine(xaml);

            XmlDocument doc = new XmlDocument();
          
            using (TextReader sr = new StringReader(xaml))
            {
                doc.Load(sr);

            }
        
            //traverse all Paragraph nodes in richtextbox xaml
            XElement xmlTree = XElement.Parse(xaml);

           
            float heightOffset = 100;
            int j = -1;

                     
            foreach (XElement node in xmlTree.Elements())
            {
                //traverse all inner Run tags

                
                foreach (XElement n in node.Elements())
                {
                    
                    Microsoft.Office.Interop.PowerPoint.Shape shape = slide.Shapes[2];
                   
                                  

                    if(n.Value==" ")
                    {
                       
                        shape = slide.Shapes[2];
                        try
                        {
                            slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 20, heightOffset, 200, 113).Fill.UserPicture(picsCollectionHelper.picsList.ElementAt(++j).getPictureUri());

                            heightOffset += 115;

                          
                        }

                        catch (Exception ex)
                        {
                            Debug.WriteLine(ex);
                        }


                    }
                    else if(n.Value=="")
                    {
                        objText.Text += "\n";
                    }
                    else
                    {

                        var textBox = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 20, heightOffset, 600, 36);
                        textBox.TextEffect.FontSize = 12;
                        textBox.TextEffect.Alignment = MsoTextEffectAlignment.msoTextEffectAlignmentWordJustify;
                        textBox.TextFrame.MarginBottom = 10;
                        textBox.TextFrame.MarginTop = 10;
                        textBox.TextFrame.TextRange.Text = n.Value;
                                            

                        int lines = n.Value.Length / 120;
                        heightOffset += 18*lines+20;

                                            
                    }
                   
                }
              
            }


            //clean all temp pics
            System.IO.DirectoryInfo di = new DirectoryInfo("C:\\temp\\Images");

            foreach (FileInfo file in di.GetFiles())
            {
                file.Delete();
            }

           
            picsCollectionHelper.picsList.Clear();
            picsCollectionHelper.nullify();
            keywords.Clear();

            pptPresentation.Save();
            
            Console.WriteLine("Stop for ppt slide debug.");

        }


    }


    public class SlideHelper
    {
        int _id;

        public SlideHelper(int id)
        {
            _id = id;
        }
        public int getId()
        {
            return _id;
        }

        public void setId(int id)
        {
            _id = id;
        }
    }

    public class SlidesCollectionHelper : IEnumerable<SlideHelper>
    {
        public List<SlideHelper> slides;
        int i;

        public SlidesCollectionHelper()
        {
            slides = new List<SlideHelper>();
            i = 0;
        }

       public void addSlide(SlideHelper slide)
        {
            ++i;
            slides.Add(slide);
            slide.setId(i);
        }


        public IEnumerator<SlideHelper> GetEnumerator()
        {
            return slides.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }
    }

    public class PictureHelper
    {
        private string _pictureUri = null;
        private int _picId;

        public PictureHelper()
        {
          
        }

        public void setPictureUri(string pictureUri)
        {
            _pictureUri = pictureUri;
        }
        public void setPictureId(int id)
        {
            _picId = id;
        }

        public string getPictureUri()
        {
            return _pictureUri;
        }

        public int getPictureId()
        {
            return _picId;
        }
    }

    public class PicsCollectionHelper : IEnumerable<PictureHelper>
    {
        public List<PictureHelper> picsList;
        int i;

        public PicsCollectionHelper()
        {
            picsList = new List<PictureHelper>();
            i = 0;
        }

        public void addPic(PictureHelper picture)
        {
            ++i;
            picsList.Add(picture);
            picture.setPictureId(i);
        }

        public void nullify()
        {
            i = 0;
        }


        public IEnumerator<PictureHelper> GetEnumerator()
        {
            return picsList.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }
    }



}

