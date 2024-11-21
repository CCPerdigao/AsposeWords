using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using Aspose.Words;
using OSAsposeWords.Structures;
using static OSAsposeWords.AsposeWords;

namespace OSAsposeWords
{
    public class AsposeWords : IAsposeWords
    {
            public static String ESCAPE_CHAR = "�";

            /// <summary>
            /// 
            /// </summary>
            /// <param name="sslicFile">Licence file</param>
            /// <param name="ssFile"></param>
            /// <param name="ssResultFile"></param>
            public AsposeResult DocToPdf(byte[] licFile, byte[] file)
            {
                byte[] ssResultFile = new byte[] { };
                string ssResultMsg = "";
                bool ssResult = false;

                try
                {
                    //trying to set the license file. if empty will get a water mark
                    if(licFile.Length >0)
                        setAsposeLicense(licFile);

                    MemoryStream m = new MemoryStream(file);
                    Document doc = new Document(m);

                    MemoryStream xmlStream = new MemoryStream();
                    doc.Save(xmlStream, Aspose.Words.SaveFormat.Pdf);
                    xmlStream.Seek(0, SeekOrigin.Begin);
                    ssResultFile = new byte[xmlStream.Length];
                    xmlStream.Read(ssResultFile, 0, (int)xmlStream.Length);
                    ssResult = true;
                    return new AsposeResult(ssResult, ssResultMsg, ssResultFile);
                }
                catch (Exception e)
                {
                    ssResult = false;
                    ssResultMsg = e.Message;
                return new AsposeResult(ssResult, ssResultMsg, new byte[] {});
                }

            } // MssDocToPdf

            /// <summary>
            /// 
            /// </summary>
            /// <param name="sslicFile">Licence file</param>
            /// <param name="ssFile"></param>
            /// <param name="ssKeywords"></param>
            /// <param name="ssValues"></param>
            /// <param name="ssMailMergeFile"></param>
            public AsposeResult MailMergeDoc(byte[] licFile, byte[] file, string keywords, string values)
            {
                byte[] ssMailMergeFile = new byte[] { };
                string ssResultMsg = "";
                Boolean ssResult = false;

                char[] delimiterChars = { ',' };

                try
                {
                    //trying to set the license file. if empty will get a water mark
                    if (licFile.Length > 0)
                        setAsposeLicense(licFile);
                    MemoryStream m = new MemoryStream(file);
                    Document doc = new Document(m);



                    String[] parsedValues = values.Split(delimiterChars);

                    String[] fields = keywords.Split(delimiterChars);

                    //Unescape comma character
                    for (int i = 0; i < parsedValues.Length - 1; i++)
                    {
                        parsedValues[i] = parsedValues[i].Replace(ESCAPE_CHAR, ",");
                    }

                    doc.MailMerge.Execute(fields, parsedValues);

                    MemoryStream xmlStream = new MemoryStream();
                    doc.Save(xmlStream, Aspose.Words.SaveFormat.Docx);
                    xmlStream.Seek(0, SeekOrigin.Begin);
                    ssMailMergeFile = new byte[xmlStream.Length];
                    xmlStream.Read(ssMailMergeFile, 0, (int)xmlStream.Length);
                    ssResult = true;
                    
                    return new AsposeResult(ssResult, ssResultMsg, ssMailMergeFile);
                }
                catch (Exception e)
                {
                    ssResult = false;
                    ssResultMsg = e.Message;
                    return new AsposeResult(ssResult, ssResultMsg, new byte[] { });
                }
            } // MssMailMergeDoc

            /// <summary>
            /// 
            /// </summary>
            /// <param name="sslicFile"></param>
            /// <param name="ssFile"></param>
            /// <param name="ssBookmark">Keyword to be replaced, if needed</param>
            /// <param name="ssImage"></param>
            /// <param name="ssResultFile"></param>
            public AsposeResult InsertImage(byte[] licFile, byte[] file, string bookmark, byte[] image)
            {
                byte[] ssResultFile = new byte[] { };
                string ssResultMsg = "";
                Boolean ssResult = false;

                MemoryStream m = new MemoryStream(file);

                try
                {
                    //trying to set the license file. if empty will get a water mark
                    if (licFile.Length > 0)
                        setAsposeLicense(licFile);

                    Document doc = new Document(m);
                    DocumentBuilder builder = new DocumentBuilder(doc);

                    MemoryStream xmlStream = new MemoryStream();
                    if (bookmark.Length > 0)
                        builder.MoveToBookmark(bookmark);
                    builder.InsertImage(image);
                    doc.Save(xmlStream, Aspose.Words.SaveFormat.Docx);

                    xmlStream.Seek(0, SeekOrigin.Begin);

                    ssResultFile = new byte[xmlStream.Length];
                    xmlStream.Read(ssResultFile, 0, (int)xmlStream.Length);

                    ssResult = true;
                    return new AsposeResult(ssResult, ssResultMsg, ssResultFile);
                }
                catch (Exception e)
                {
                    ssResult = false;
                    ssResultMsg = e.Message;
                    return new AsposeResult(ssResult, ssResultMsg, new byte[] { });
                }


            } // MssInsertImage

            private void setAsposeLicense(byte[] sslicFile)
            {
                Aspose.Words.License WordsLicense = new Aspose.Words.License();
                Stream streamLic = new MemoryStream(sslicFile);
                WordsLicense.SetLicense(streamLic);
            }


        } // CssAsposeWords

    } // OutSystems.NssAsposeWords
