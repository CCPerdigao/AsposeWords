using OSAsposeWords;
using OSAsposeWords.Structures;

namespace AsposeWordsUnitTesting
{
    [TestClass]
    public class EmptyLicenseTest
    {
        [TestMethod]
        public void DocToPDFTest()
        {
            AsposeWords aw = new AsposeWords();
            AsposeResult docRes, docxRes;
            byte[] doc = System.IO.File.ReadAllBytes("./Resources/OutSystemsDoc.Doc"); 
            docRes = aw.DocToPdf(new byte[] { }, doc);
            Assert.IsTrue(docRes.success);
            System.IO.File.WriteAllBytes("./Resources/OutSystemsDoc.pdf", docRes.document);
            

            byte[] docx = System.IO.File.ReadAllBytes("./Resources/OutSystemsDocX.Docx"); 
            docxRes = aw.DocToPdf(new byte[] { }, docx);
            Assert.IsTrue(docxRes.success);
            System.IO.File.WriteAllBytes("./Resources/OutSystemsDocX.pdf", docxRes.document);

        }

        [TestMethod]
        public void DocMergeTest()
        {
            AsposeWords aw = new AsposeWords();
            AsposeResult docRes;
            byte[] doc = System.IO.File.ReadAllBytes("./Resources/OutSystemsMailMerge.Dotx"); ;
            
            docRes = aw.MailMergeDoc(new byte[] { }, doc, "NAME,COMPANY","CARLOS PERDIGÃO,OUTSYSTEMS");
            Assert.IsTrue(docRes.success);
            System.IO.File.WriteAllBytes("./Resources/OutSystemsMerge.docx", docRes.document);
        }

        [TestMethod]
        public void InsertImageTest()
        {

            AsposeWords aw = new AsposeWords();
            AsposeResult docRes;
            byte[] doc = System.IO.File.ReadAllBytes("./Resources/OutSystemsImage.Docx"); 
            byte[] pic = System.IO.File.ReadAllBytes("./Resources/odc.png"); ;
            docRes = aw.InsertImage(new byte[] { }, doc,"SecondPic",pic);
            Assert.IsTrue(docRes.success);
            System.IO.File.WriteAllBytes("./Resources/OutSystemsPic.docx", docRes.document);
        }
    }
}