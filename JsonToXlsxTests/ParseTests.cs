using System;
using System.IO;
using System.Text;
using JsonToXlsx.NETCore;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace JsonToXlsxTests
{
    [TestClass]
    public class ParseTests
    {
        [TestInitialize]
        public void TestInitialize()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }
        
        //[TestMethod]
        public void JsonToXlsxTest()
        {
            const string sourceFilePath = @"path";
            string resultPath = Path.Combine(
                Path.GetDirectoryName(sourceFilePath),
                Path.GetFileNameWithoutExtension(sourceFilePath) + DateTime.Now.ToString("dd.MM.yyyy_HH.mm.ss") + ".xlsx");
            string json = File.ReadAllText(sourceFilePath, Encoding.UTF8);
            JsonToXlsxConverter.SaveAs(json, resultPath);
        }

        //[TestMethod]
        public void XlsxToJsonTest()
        {
            string filePath = @"path";

            JArray jObject = JsonToXlsxConverter.XlsxToJson(filePath);
            
            File.WriteAllText(
                Path.Combine(
                    Path.GetDirectoryName(filePath),
                    Path.GetFileNameWithoutExtension(filePath) + ".json"),
                JsonConvert.SerializeObject(
                    jObject,
                    Formatting.Indented));
        }
    }
}