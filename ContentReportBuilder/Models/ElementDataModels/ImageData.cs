using System.IO;
using System.Net;
using System.Threading.Tasks;

namespace ContentReportBuilder.Models.ElementDataModels
{
    /// <summary>
    /// Image element model data
    /// </summary>
    public abstract class ImageDataBase
    {
        public ImageDataBase()
        {

        }
        public ImageDataBase(int height = 60, int width = 40)
        {
            Width = width;
            Height = height;
        }

        public int Height { get; set; }
        public int Width { get; set; }

        /// <summary>
        /// Returns a stream of the image content
        /// </summary>
        /// <returns></returns>
        public abstract Stream GetData();
    }

    public class WebImageData: ImageDataBase
    {
        public WebImageData(): base()
        {

        }
        public WebImageData(string url, int height = 60, int width = 40) : base(height, width)
        {
            URL = url;
        }
        public string URL { get; set; }

        /// <summary>
        /// Uses the url provided in the constructor to find the image on the web and return it as a stream.
        /// If the http request fails (response not in the 200s) null is returned.
        /// </summary>
        /// <returns>Stream containing the image data, or null if the http request fails</returns>
        public override Stream GetData()
        {
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(URL);
            req.UseDefaultCredentials = true;
            req.PreAuthenticate = true;
            req.Credentials = CredentialCache.DefaultCredentials;
            var resp = req.GetResponse();

            if((int)(resp as HttpWebResponse).StatusCode >= 200 || (int)(resp as HttpWebResponse).StatusCode < 300)
            {
                return (resp as HttpWebResponse).GetResponseStream();
            }

            return null;
        }
    }
    /// <summary>
    /// Uses the file path provided in the constructor to find the image on the local file system and return it as a stream.
    /// If the image cannot be found null is returned.
    /// </summary>
    /// <returns>Stream containing the image data, or null if the image cannot be found</returns>
    public class LocalImageData : ImageDataBase
    {
        public LocalImageData() : base()
        {

        }
        public LocalImageData(string filepath, int height = 60, int width = 40) : base(height, width)
        {
            FilePath = filepath;
        }
        public string FilePath { get; set; }
        public override Stream GetData()
        {
            if (File.Exists(FilePath))
            {
                return File.OpenRead(FilePath);
            }

            return null;
        }
    }
}
