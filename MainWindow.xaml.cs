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
        List<String> tokens = new List<String>();


        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
        }

        private void titleChangedEventHandler(object sender, TextChangedEventArgs args)
        {
           /* string title = titleArea.Text;
            string[] words = title.Split(' ');
           
            foreach (var el in words)
            {
                keywords.Add(el);
                //Console.WriteLine(el);
            }*/
                       
        }

        private void getTitleKeywords()
        {
            string title = titleArea.Text;
            string[] words = title.Split(' ');

            foreach (var el in words)
            {
                if(!keywords.Contains(el))
                    keywords.Add(el);
                //Console.WriteLine(el);
            }
        }
       
        private void textChangedEventHandler(object sender, TextChangedEventArgs args)
        {/*
            tokens.Clear();
            string temp = "";
           
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

            if (coll.Count>0)
            {
               *//* for (int i = 0; i < coll.Count; i++)
                {
                    result = coll[i].Groups[2].Value;
                    tmp.Add(result);
                }*//*



                temp = coll[coll.Count - 1].Groups[2].Value;
                
            }

            string[] token = temp.Split(' ');
          
            if (token.Length > 1)
            {
                foreach(var t in token)
                {
                    tokens.Add(t);
                }
            }       
            
            
            foreach (var el in tokens)
            {
                
                    keywords.Add(el);
                    Console.WriteLine(el);       
            }*/

  /*          distinct = keywords.Distinct().ToList();*/


        }

        private void textSelectionChangedEventHandler(object sender, DependencyPropertyChangedEventArgs e)
        {
/*            List<String> tmp = new List<String>();

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
                if(keywords.Contains(el)==false)
                {
                    keywords.Add(el);
                }
                    
               
            }

            foreach (var el in keywords)
                Console.WriteLine(el);*/

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

                    //create the search tasks to be executed
                    /* var tasks = new[]{
                                         GetAsync("flower", 4),
                                         GetAsync("cat", 4),
                                         GetAsync("dog", 4),
                                     };*/


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
                        foreach (var l in imgUrls)
                        {
                            urls.Add(l);
                            Console.WriteLine(l);
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
               + "&imgSize=SMALL" +
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
    }
}
