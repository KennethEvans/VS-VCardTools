using System;
using System.Drawing;
using System.Drawing.Imaging;
using EWSoftware.PDI.Objects;
using EWSoftware.PDI.Parser;
using EWSoftware.PDI.Properties;

namespace VCardExtractPhotos {
    class Program {
        public static string LF = System.Environment.NewLine;
        private static VCardCollection vCards;
        private static string inutFileName = @"C:\Scratch\Contacts\Contacts-S7-2016-12-29.vcf";
        //private static string inutFileName = @"C:\Scratch\Contacts\Contacts-S7-2016-12-29.KennethEvans.vcf";
        private static string outputDirName = @"C:\Scratch\Contacts\Photos\";

        private static void saveImages() {
            try {
                vCards = VCardParser.ParseFromFile(inutFileName);
            } catch (Exception ex) {
                string error = String.Format("Unable to load vCards:\n{0}",
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
                    saveImage(vCard);
                }
            }
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
            saveImages();
        }
    }
}
