using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXMLMailMerge.Core.OpenXMLDocument.Dictionary;
using System.IO;
using System.Data;
using OpenXMLMailMerge;

namespace OpenXMLMailMergeTest
{
    [TestClass]
    public class LoadTest
    {
        const string bindFile = @".\doc\BIND.docx";// BIND.docx";
        const string imageFile = @".\doc\image.png";

        private OXMailMerge GetOXObject()
        {
            var imageData = File.ReadAllBytes(imageFile);

            //Just testing the byte array integrity
            //var sw = new StreamWriter(@"D:\vrl\backup\Documents\vrl\lab\TestLab\OpenXMLMailMerge\OpenXMLMailMerge\OpenXMLMailMergeTest\doc\image2.png");
            //sw.BaseStream.Write(imageData, 0, imageData.Length);
            //sw.Close();

            var dt = new DataTable();
            dt.Columns.Add(new DataColumn("COL1", typeof(string)));
            dt.Columns.Add(new DataColumn("COL2", typeof(int)));
            dt.Columns.Add(new DataColumn("COL3", typeof(DateTime)));
            dt.Columns.Add(new DataColumn("COL4", typeof(decimal)));

            for (int i = 0; i < 1; i++)
            {
                var row = dt.NewRow();
                row["COL1"] = $"ROW {i.ToString()}";
                row["COL2"] = i;
                row["COL3"] = DateTime.Now;
                row["COL4"] = Convert.ToDecimal(i);
                dt.Rows.Add(row);
            }

            var openXMLMailMerge = new OpenXMLMailMerge.OXMailMerge(DocumentType.DOC);
            openXMLMailMerge.Document.AddToDictionary(MailMergeDataTypeEnum.Regex, "HEADER", "Vagner");
            openXMLMailMerge.Document.AddToDictionary(MailMergeDataTypeEnum.Regex, "FOOTER", "Lucas");
            openXMLMailMerge.Document.AddToDictionary(MailMergeDataTypeEnum.Regex, "BIND", "Reis");
            openXMLMailMerge.Document.AddToDictionary(MailMergeDataTypeEnum.Image, "IMG", imageData);
            openXMLMailMerge.Document.AddToDictionary(MailMergeDataTypeEnum.Table, "TABLE", dt);

            return openXMLMailMerge;
        }

        [TestMethod]
        public void LoadFromBytesTest()
        {
            var openXMLMailMerge = GetOXObject();
            var fileBytes = File.ReadAllBytes(bindFile);
            byte[] data = null;

            try
            {
                var tempPath = @".\doc\";
                openXMLMailMerge.Document.LoadFromBytes(fileBytes, tempPath);
                openXMLMailMerge.Process();
                openXMLMailMerge.Document.SaveToFile(tempPath + "vrl.docx");
                data = openXMLMailMerge.Document.GetBytes();
                File.WriteAllBytes(tempPath + "out.docx", data);
                openXMLMailMerge.Terminate();
            }
            catch (Exception)
            {
                throw;
            }

            Assert.AreEqual(true, data != null && data.Length > 0);
        }

        [TestMethod]
        public void LoadFileTest()
        {
            var result = false;

            try
            {
                var openXMLMailMerge = GetOXObject();
                openXMLMailMerge.Document.LoadFromFile(bindFile);

                openXMLMailMerge.Process();
                openXMLMailMerge.Document.SaveToFile();
                openXMLMailMerge.Terminate();

                result = true;
            }
            catch (Exception)
            {
                throw;
            }

            Assert.AreEqual(true, result);
        }
    }
}
