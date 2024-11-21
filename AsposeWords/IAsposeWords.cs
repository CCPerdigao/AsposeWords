using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OutSystems.ExternalLibraries.SDK;
using OSAsposeWords.Structures;



namespace OSAsposeWords
{

    [OSInterface(Description = "Provides actions to manipulate .DOC documents using the AsposeLibrary", IconResourceName = "AsposeWords.Resources.Aspose.jpeg")]
    public interface IAsposeWords
    {

        /// <summary>
        /// This action Converts a Word documents (.doc) into a pdf file.
        /// </summary>
        /// <param name="licFile">Licence file provided by Aspose</param>
        /// <param name="file">.Doc file to be converted to PDF</param>
        public AsposeResult DocToPdf(byte[] licFile, byte[] file);

        /// <summary>
        /// This action replaces keywords with the respective value
        /// </summary>
        /// <param name="licFile">Licence file provided by Aspose</param>
        /// <param name="file">.Doc file to be processed</param>
        /// <param name="keywords">String with the values to be replaced separated by ,</param>
        /// <param name="values">String with the values to replace with separated by ,</param>
        public AsposeResult MailMergeDoc(byte[] licFile, byte[] file, string keywords, string values);

        /// <summary>
        /// Insert image into a bookmark in the .Doc document
        /// </summary>
        /// <param name="licFile">Licence file provided by Aspose</param>
        /// <param name="file">.Doc file to be processed</param>
        /// <param name="bookmark">Keyword to be replaced, if needed</param>
        /// <param name="image">Binary with the image to insert into the document</param>
        public AsposeResult InsertImage(byte[] licFile, byte[] file, string bookmark, byte[] image);

    } // IssAsposeWords

} // OutSystems.NssAsposeWords