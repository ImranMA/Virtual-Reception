namespace VirtualReception
{
    public class LUISResponse
    {
        public string query { get; set; }
        public lIntent[] intents { get; set; }
        public lEntity[] entities { get; set; }
        public sentiments sentimentAnalysis { get; set; }
    }

    public class sentiments
    {
        public string label { get; set; }
        public float score { get; set; }

    }    

    public class lIntent
    {
        public string intent { get; set; }
        public float score { get; set; }
    }

    public class lEntity
    {
        public string entity { get; set; }
        public string type { get; set; }
        public int startIndex { get; set; }
        public int endIndex { get; set; }
        public float score { get; set; }
    }
}
