namespace updateTemplateWeb.Models
{
    public class ModelData
    {
        public string ORIGEN { get; set; }
        public string DESTINATION { get; set; }
        public string PRODUCT { get; set; }
        public string RATE { get; set; }
    }

    public class AlertMessageViewModel
    {
        public string MessageType { get; set; }
        public string Message { get; set; }
    }
}