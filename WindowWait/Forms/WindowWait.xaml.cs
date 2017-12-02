using GeneralClasses;
using SoftwareClasses;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Xml;

namespace WindowWait
{
    /// <summary>
    /// Class MainWindow_Wait from the namespace WindowWait
    /// </summary>
    public partial class MainWindow_Wait : Window
    {

        #region Initialization

        #region Variables

        private Thread StatusThread = null;

        private MainWindow_Wait Popup = null;

        /// <summary>
        /// Global handlers for common objects
        /// </summary>
        private static Handlers m_Global_Handler = null;

        /// <summary>
        /// Principal message
        /// </summary>
        private string m_PrincipalMessage = "";

        /// <summary>
        /// Secondary message
        /// </summary>
        private string m_SecondaryMessage = "";

        #endregion

        #region Functions

        /// <summary>
        /// Default constructor for the wait main window
        /// </summary>
        public MainWindow_Wait()
        {
            //Nothing to do
        }

        /// <summary>
        /// Constructor for the wait main window
        /// </summary>
        private MainWindow_Wait(Handlers _Global_Handler, string _PrincipalMessage, string _SecondaryMessage = "")
        {
            try
            {
                //Intialize the components
                InitializeComponent();

                //Initialize variables,
                m_Global_Handler = _Global_Handler;
                m_PrincipalMessage = m_Global_Handler.Resources_Handler.Get_Resources(_PrincipalMessage);
                m_SecondaryMessage = m_Global_Handler.Resources_Handler.Get_Resources(_SecondaryMessage);

                //Define content
                if (m_Global_Handler != null)
                {
                    Define_Content();
                }
            }
            catch (Exception exception)
            {
                m_Global_Handler.Log_Handler.WriteException(System.Reflection.MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        /// <summary>
        /// Define each content for labels, button, text in the form
        /// </summary>
        private void Define_Content()
        {
            Lbl_PrincipalMessage.Content = new TextBlock() { Text = m_PrincipalMessage, TextWrapping = TextWrapping.Wrap };
            Lbl_SecondaryMessage.Content = new TextBlock() { Text = m_SecondaryMessage, TextWrapping = TextWrapping.Wrap };
        }

        #endregion

        #endregion

        #region Functions

        /// <summary>
        /// Open the window through a new thread
        /// </summary>
        public void Start(Handlers _Global_Handler, string _PrincipalMessage, string _SecondaryMessage = "")
        {
            //create the thread with its ThreadStart method
            this.StatusThread = new Thread(() =>
            {
                try
                {
                    this.Popup = new MainWindow_Wait(_Global_Handler, _PrincipalMessage, _SecondaryMessage);
                    this.Popup.Show();
                    this.Popup.Closed += (lsender, le) =>
                    {
                        //when the window closes, close the thread invoking the shutdown of the dispatcher
                        this.Popup.Dispatcher.InvokeShutdown();
                        this.Popup = null;
                        this.StatusThread = null;
                    };

                    //this call is needed so the thread remains open until the dispatcher is closed
                    System.Windows.Threading.Dispatcher.Run();
                }
                catch (Exception exception)
                {
                    m_Global_Handler.Log_Handler.WriteException(System.Reflection.MethodBase.GetCurrentMethod().Name, exception);
                    return;
                }
            });

            //run the thread in STA mode to make it work correctly
            this.StatusThread.SetApartmentState(ApartmentState.STA);
            this.StatusThread.Priority = ThreadPriority.Normal;
            this.StatusThread.Start();
        }

        /// <summary>
        /// Close the window and end the associated thread
        /// </summary>
        public void Stop()
        {
            try
            {
                if (this.Popup != null)
                {
                    //need to use the dispatcher to call the Close method, because the window is created in another thread, and this method is called by the main thread
                    this.Popup.Dispatcher.BeginInvoke(new Action(() =>
                    {
                        this.Popup.Close();
                    }));
                }
            }
            catch (Exception exception)
            {
                m_Global_Handler.Log_Handler.WriteException(System.Reflection.MethodBase.GetCurrentMethod().Name, exception);
                return;
            }
        }

        #endregion

    }
}
