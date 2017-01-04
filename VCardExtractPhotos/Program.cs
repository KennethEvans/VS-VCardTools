using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using EWSoftware.PDI.Objects;
using EWSoftware.PDI.Parser;
using EWSoftware.PDI.Properties;

namespace VCardExtractPhotos {
    class Program {
        public static string LF = System.Environment.NewLine;
        /// <summary>
        /// The input file name for saveImages and the file with images for
        /// transferImages.
        /// </summary>
        private static string inputFileName = @"C:\Scratch\Contacts\Contacts-S7-2016-12-29.vcf";
        //private static string inutFileName = @"C:\Scratch\Contacts\Contacts-S7-2016-12-29.KennethEvans.vcf";
        /// <summary>
        /// The file without images for transferImages.
        /// </summary>
        private static string inputFileName2 = @"C:\Scratch\Contacts\google.contacts.MyContacts.010217.vcf";
        /// <summary>
        /// The out directory for saveImages.  The image files go there.
        /// </summary>
        private static string outputDirName = @"C:\Scratch\Contacts\Photos\";
        /// <summary>
        /// The output directory for the converted file for transferImages.  The
        /// converted file goes there.
        /// </summary>
        private static string outputDirName2 = @"C:\Scratch\Contacts\";

        /// <summary>
        /// Transfers images in inputFileName to vCards in inputFileName2
        /// having the same FormattedName and saves the result in outputDirName2.
        /// </summary>
        private static void transferImages() {
            VCardCollection vCards, vCards2;
            try {
                vCards = VCardParser.ParseFromFile(inputFileName);
            } catch (Exception ex) {
                string error = String.Format("Unable to load vCards from "
                    + inputFileName + ":\n{0}",
                    ex.Message);
                if (ex.InnerException != null) {
                    error += ex.InnerException.Message + "\n";
                    if (ex.InnerException.InnerException != null)
                        error += ex.InnerException.InnerException.Message;
                }
                System.Diagnostics.Debug.Write(ex);
                Utils.excMsg(error, ex);
                return;
            }
            Console.WriteLine("Read " + inputFileName);
            Console.WriteLine("There are " + vCards.Count + " vCards");
            try {
                vCards2 = VCardParser.ParseFromFile(inputFileName2);
            } catch (Exception ex) {
                string error = String.Format("Unable to load vCards from "
                    + inputFileName2 + ":\n{0}",
                    ex.Message);
                if (ex.InnerException != null) {
                    error += ex.InnerException.Message + "\n";
                    if (ex.InnerException.InnerException != null)
                        error += ex.InnerException.InnerException.Message;
                }
                System.Diagnostics.Debug.Write(ex);
                Utils.excMsg(error, ex);
                return;
            }
            Console.WriteLine("Read " + inputFileName2);
            Console.WriteLine("There are " + vCards2.Count + " vCards");

            // Loop over the second file
            foreach (VCard vCard2 in vCards2) {
                foreach (VCard vCard in vCards) {
                    if (vCard2.Name.FormattedName
                        .Equals(vCard.Name.FormattedName)) {
                        // The names match
                        if (vCard.Photo.Value != null &&
                            vCard.Photo.ValueLocation == ValLocValue.Binary) {
                            // There is an image in the first one
                            byte[] imageBytes = vCard.Photo.GetImageBytes();
                            vCard2.Photo.SetImageBytes(imageBytes);
                            vCard2.Photo.ImageType = vCard.Photo.ImageType;
                        }
                        break;
                    }
                } // foreach vCard1
            }  // foreach vCard2

            string name = Path.GetFileName(inputFileName2);
            string prefix = Path.GetFileNameWithoutExtension(name);
            string ext = Path.GetExtension(inputFileName2);
            string outputFileName2 = Path.Combine(outputDirName2,
                name + ".Combined" + ext);

            using (var sw = new StreamWriter(outputFileName2, false,
                PDIParser.DefaultEncoding)) {
                vCards.WriteToStream(sw);
            }
            Console.WriteLine("Wrote " + outputFileName2);
        }

        /// <summary>
        /// Finds images in inputFileName and saves them as PNG files in
        /// outputDirName2.
        /// </summary>
        private static void saveImages() {
            VCardCollection vCards;
            try {
                vCards = VCardParser.ParseFromFile(inputFileName);
            } catch (Exception ex) {
                string error = String.Format("Unable to load vCards from "
                    + inputFileName + ":\n{0}",
                      ex.Message);
                if (ex.InnerException != null) {
                    error += ex.InnerException.Message + "\n";
                    if (ex.InnerException.InnerException != null)
                        error += ex.InnerException.InnerException.Message;
                }
                System.Diagnostics.Debug.Write(ex);
                Utils.excMsg(error, ex);
                return;
            }

            // Loop over cards
            string imageFileName;
            bool hasImage;
            string imageType;
            string info = "";
            int nPhotos = 0;
            foreach (VCard vCard in vCards) {
                hasImage = false;
                imageFileName = null;
                info += vCard.Name.FormattedName + ": ";
                if (vCard.Photo.Value != null) {
                    hasImage = true;
                    imageType = vCard.Photo.ImageType;
                    if (vCard.Photo.ValueLocation == ValLocValue.Binary) {
                        info += "Binary Image: " + imageType;
                    } else {
                        imageFileName = vCard.Photo.Value;
                        info += "Image: " + imageType + " " + imageFileName;
                    }
                } else {
                    hasImage = false;
                    info += "No Image";
                }
                info += LF;
                Console.Write(info);
                if (hasImage) {
                    nPhotos++;
                    saveImage(vCard);
                }
            }
            Console.WriteLine("There are " + vCards.Count + " vCards in "
                + inputFileName);
            Console.WriteLine("Wrote " + nPhotos + " photos to "
                + outputDirName);
        }

        static void saveImage(VCard vCard) {
            PhotoProperty photo = vCard.Photo;
            if (photo.ValueLocation != ValLocValue.Binary) {
                return;
            }
            byte[] imageBytes = photo.GetImageBytes();
            Bitmap bm;
            try {
                bm = new Bitmap(new System.IO.MemoryStream(imageBytes));
            } catch (Exception ex) {
                Utils.excMsg("Failed to create Bitmap", ex);
                return;
            }
            try {
                string fileName = outputDirName + vCard.Name.FormattedName
                    + ".png";
                bm.Save(fileName, ImageFormat.Png);
                Console.Write("  Saved photo as " + fileName + LF);
            } catch (Exception ex) {
                Utils.excMsg("Failed to save Bitmap", ex);
            }
        }

        static void Main(string[] args) {
            //saveImages();
            transferImages();
        }
    }
}
