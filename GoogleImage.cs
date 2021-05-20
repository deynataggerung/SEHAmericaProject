using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace WindowsFormsApplication1
{
    internal class GoogleImageItems
    {
        [JsonProperty("items")]
        public List<GoogleImage> items { get; set; }
    }
    internal class GoogleImage
    {
        public string title { get; set; }
        public string link { get; set; }
        public Image image { get; set; }
    }
    internal class Image
    {
        public string height { get; set; }
        public string width { get; set; }
        public string thumbnailLink { get; set; }
    }
}