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
        Slides slideCollection = new Slides();
       
        byte[] imageBytes = null;
        

        public MainWindow()
        {
            InitializeComponent();
           
            mainRTB.AllowDrop = true;
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

            TextRange allText = new TextRange(mainRTB.Document.ContentStart, mainRTB.Document.ContentEnd);

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
                        alltasks[i] = GetAsync(keywords[i], 2);
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
                imageBytes = webClient.DownloadData(url);

                DataObject data = new DataObject();

                data.SetData(imageBytes);

                DragDrop.DoDragDrop(this, data, DragDropEffects.Copy | DragDropEffects.Move);

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

                    // Set Effects to notify the drag source what effect
                    // the drag-and-drop operation had.
                    // (Copy if CTRL is pressed; otherwise, move.)
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
                      BitmapSizeOptions.FromWidthAndHeight(bmp.Width, bmp.Height));
                    
                    imgControl.Source = bs;

                    mainRTB.Document.Blocks.Add(paragraph);

                    bmp.Dispose();

                    Debug.WriteLine("Stop here.");

                    // Set Effects to notify the drag source what effect
                    // the drag-and-drop operation had.
                    // (Copy if CTRL is pressed; otherwise, move.)
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

            /* string filePath = @"test.rtf";
             FileStream fs = new FileStream(filePath, FileMode.Create, FileAccess.Write);
             allText.Save(fs, DataFormats.Rtf);
             fs.Close();*/

            //create slide and add to collection of slides
            Slide slide = new Slide(memstream);
            slideCollection.addSlide(slide);

            mainRTB.Document.Blocks.Clear();
            imagesStackPanel.Children.Clear();

            Console.WriteLine("stop.");
        }

        private MemoryStream saveToMemStream()
        {
            TextRange allText = new TextRange(mainRTB.Document.ContentStart, mainRTB.Document.ContentEnd);

            MemoryStream memstream = new MemoryStream();
            allText.Save(memstream, DataFormats.Rtf);

            /*if (memstream != null)
            {
                memstream.Close();
            }*/

            return memstream;
        }

        private void SaveAll(object sender, RoutedEventArgs e)
        {
            if (slideCollection == null) return;

           
            foreach (var slide in slideCollection)
            {
                saveSlide(slide);
            }
        }


        private void saveSlide(Slide slide)
        {
            string filePath = @"test" + slide.getId() + ".rtf";
            FileStream fs = new FileStream(filePath, FileMode.Create, FileAccess.Write);
            //memoryStream.WriteTo(fileStream);
            slide.getSlide().WriteTo(fs);

            fs.Close();
        }


    }


    public class Slide
    {
        MemoryStream _slide;
        int _id;

        public Slide(MemoryStream ms)
        {
            _slide = new MemoryStream();
            _slide = ms;
        }

        public void setId(int id)
        {
            _id = id;
        }

        public int getId()
        {
            return _id;
        }

        public MemoryStream getSlide()
        {
            return _slide;
        }
    }

    public class Slides : IEnumerable<Slide>
    {
        public List<Slide> slides;
        int i;

        public Slides()
        {
            slides = new List<Slide>();
            i = 0;
        }

       public void addSlide(Slide slide)
        {
            ++i;
            slides.Add(slide);
            slide.setId(i);
        }


        //TODO: manage nulls
        public IEnumerator<Slide> GetEnumerator()
        {
            return slides.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }
    }

}

