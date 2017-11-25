using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;

namespace GeneralClasses
{

    /// <summary>
    /// Class Images containing all the functions for handling images
    /// </summary>
    public class Images
    {

        #region Variables

        /// <summary>
        /// Error handler 
        /// </summary>
        private static Log m_Error_Handler = new Log();

        #endregion

        #region Functions

        /// <summary>
        /// Functions
        /// Convert a bitmap to a bitmap image
        /// </summary>
        public BitmapImage Convert_ToBitmapImage(Bitmap bitmap)
        {
            using (var memory = new MemoryStream())
            {
                bitmap.Save(memory, ImageFormat.Png);
                memory.Position = 0;

                var bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.StreamSource = memory;
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.EndInit();

                return bitmapImage;
            }
        }

        /// <summary>
        /// Convert a <see cref="System.Drawing.Bitmap"/> into a WPF <see cref="BitmapSource"/>.
        /// </summary>
        /// <remarks>Uses GDI to do the conversion. Hence the call to the marshalled DeleteObject.
        /// </remarks>
        /// <param name="_Source">The source bitmap.</param>
        /// <returns>A BitmapSource</returns>
        public BitmapSource Convert_ToBitmapSource(Bitmap _Source)
        {
            BitmapSource bitSrc = null;

            var hBitmap = _Source.GetHbitmap();

            try
            {
                bitSrc = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    hBitmap,
                    IntPtr.Zero,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());
            }
            catch (Win32Exception)
            {
                bitSrc = null;
            }
            catch (Exception e)
            {
                m_Error_Handler.WriteException(System.Reflection.MethodBase.GetCurrentMethod().Name, e);
                bitSrc = null;
            }

            return bitSrc;
        }

        /// <summary>
        /// Resize a <see cref="System.Drawing.Bitmap"/> to a new Size
        /// </summary>
        /// <param name="_ImgToResize">The source bitmap.</param>
        /// <param name="_Size">The final size.</param>
        /// <returns>The resized image</returns>
        public Image ResizeImage(Image _ImgToResize, System.Drawing.Size _Size)
        {
            return (Image)(new Bitmap(_ImgToResize, _Size));
        }

        /// <summary>
        /// Resize a <see cref="System.Drawing.Bitmap"/> to a new Size
        /// </summary>
        /// <param name="_ImgToResize">The source bitmap</param>
        /// <param name="_Width">The final width</param>
        /// <param name="_Height">The final height</param>
        /// <returns>The resized image</returns>
        public Image ResizeImage(Image _ImgToResize, int _Width, int _Height)
        {
            System.Drawing.Size finalSize = new System.Drawing.Size(_Width, _Height);
            return (Image)(new Bitmap(_ImgToResize, finalSize));
        }

        /// <summary>
        /// Test if a file is an image
        /// </summary>
        /// <param name="_Filename">File name of the file to test</param>
        /// <param name="_GlobalHandler">Global handler to create the message box</param>
        /// <returns>True if the file is an image, false otherwise</returns>
        public bool Is_ImageFromFilename(string _Filename, Handlers _GlobalHandler)
        {
            try
            {
                //Get the extension
                string fileExtension = System.IO.Path.GetExtension(_Filename);

                //Test the extension
                if (fileExtension == null || fileExtension == "")
                {
                    return false;
                }

                //Test if the extension corresponds to an image
                fileExtension = fileExtension.ToLower();
                if (fileExtension == ".jpg" || fileExtension == ".jpeg" || fileExtension == ".bmp" || fileExtension == ".png" || fileExtension == ".gif")
                {
                    return true;
                }
                else
                {
                    System.Windows.MessageBox.Show(_GlobalHandler.Resources_Handler.Get_Resources("ImageExtensionFalse"),
                        _GlobalHandler.Resources_Handler.Get_Resources("WrongFile"),
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }
            }
            catch (Exception e)
            {
                m_Error_Handler.WriteException(System.Reflection.MethodBase.GetCurrentMethod().Name, e);
                return false;
            }
        }

        /// <summary>
        /// Functions
        /// Load a bitmap image from a file path
        /// </summary>
        public BitmapImage Load_BitmapImage(string _Path)
        {
            //Open file in read only mode
            using (FileStream stream = new FileStream(_Path, FileMode.Open, FileAccess.Read))
            //Get a binary reader for the file stream
            using (BinaryReader reader = new BinaryReader(stream))
            {
                //Copy the content of the file into a memory stream
                var memoryStream = new MemoryStream(reader.ReadBytes((int)stream.Length));
                //Make a new Bitmap object the owner of the MemoryStream
                Bitmap bmp = new Bitmap(memoryStream);
                //Convert to bitmap image
                return Convert_ToBitmapImage(bmp);
            }
        }

        /// <summary>
        /// Functions
        /// Load a bitmap from a file path
        /// </summary>
        public Bitmap Load_Bitmap(string _Path)
        {
            //Open file in read only mode
            using (FileStream stream = new FileStream(_Path, FileMode.Open, FileAccess.Read))
            //Get a binary reader for the file stream
            using (BinaryReader reader = new BinaryReader(stream))
            {
                //Copy the content of the file into a memory stream
                var memoryStream = new MemoryStream(reader.ReadBytes((int)stream.Length));
                //Make a new Bitmap object the owner of the MemoryStream
                Bitmap bmp = new Bitmap(memoryStream);
                return bmp;
            }
        }

        #endregion

    }
}
