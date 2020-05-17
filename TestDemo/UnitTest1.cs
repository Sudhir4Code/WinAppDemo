using System;
using System.Diagnostics;
using System.Net;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using System.Collections.Generic;
using MySqlX.XDevAPI;

namespace TestDemo
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestVmaxROVPlayer()
        {
            AppiumOptions appOptions = new AppiumOptions();

            appOptions.AddAdditionalCapability("app", @"C:\Program Files\VMAX Technologies, Inc\VMAX ROV Player x64\bin\VmPlayer2.exe");
            appOptions.AddAdditionalCapability("ms:waitForAppLaunch", "15");
            WindowsDriver<WindowsElement> ROVPLayerSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appOptions);

            Console.WriteLine("WindowsDriver is initialized");
            Console.WriteLine("Application Title: " + ROVPLayerSession.Title);

            //ROVPLayerSession.FindElementByXPath("//MenuItem[@Name='File']").Click();
            ROVPLayerSession.FindElementByXPath($"//MenuItem[starts-with(@Name, \"File\"]").Click();

            ROVPLayerSession.FindElementByAccessibilityId("18124").Click();


            

            ROVPLayerSession.FindElementByName("File").Click();


            Assert.IsFalse(ROVPLayerSession.FindElementByAccessibilityId("1025").Enabled);

            Console.WriteLine("Water Visibility (meters) spinner control is disabled");

            ROVPLayerSession.FindElementByName("Detritus").Click();
            Thread.Sleep(900);
            Console.WriteLine("Detritus button is clicked");

            Assert.IsTrue(ROVPLayerSession.FindElementByAccessibilityId("1025").Enabled);

            Console.WriteLine("Water Visibility (meters) spinner control is enabled");


            ROVPLayerSession.Close();
            ROVPLayerSession.Quit();
        }


        [TestMethod]
        public void TestVmaxEditor()
        {
            AppiumOptions appOptions = new AppiumOptions();                        
            appOptions.AddAdditionalCapability("app", @"C:\Program Files\VMAX Technologies, Inc\VMAX Editor x64\bin\VmEditor2.exe");
            appOptions.AddAdditionalCapability("ms:waitForAppLaunch", "25"); appOptions.AddAdditionalCapability("appWorkingDir", @"C:\Program Files\VMAX Technologies, Inc\VMAX Editor x64\")
           ;
            
            WindowsDriver <WindowsElement> VMAXEditorSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appOptions);
            
            Console.WriteLine("WindowsDriver is initialized");
            Console.WriteLine(VMAXEditorSession.Title);

            VMAXEditorSession.Close();
            VMAXEditorSession.Quit();                        
        }

        [TestMethod]
        public void SoapUITest()
        {
            AppiumOptions appOptions = new AppiumOptions();         
            appOptions.AddAdditionalCapability("app", @"C:\Program Files (x86)\SmartBear\SoapUI-5.5.0\bin\SoapUI-5.5.0.exe");
            appOptions.AddAdditionalCapability("ms:waitForAppLaunch", "15");
            WindowsDriver<WindowsElement> SOAPUISession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appOptions);
   
            Console.WriteLine("WindowsDriver is initialized");
            Console.WriteLine("Application Title:" + SOAPUISession.Title);
            AppiumWebElement EndPointExpClose = SOAPUISession.FindElementByName("Endpoint Explorer").FindElementByName("Close");
            EndPointExpClose.Click();
            Console.WriteLine("Endpoint explorer window is closed");

            SOAPUISession.Manage().Window.Maximize();
            Console.WriteLine("SOAPUI tool window is maximized");

            SOAPUISession.FindElementByName("Close").Click();
            Console.WriteLine("Close button is clicked");

            SOAPUISession.FindElementByName("Question").SendKeys(Keys.Enter);

            Console.WriteLine("SOAPUI application is closed");

            SOAPUISession.Quit();              
        }

        [TestMethod]
        public void AlreadyRunningApp()
        {
            AppiumOptions appOptions = new AppiumOptions();
            appOptions.AddAdditionalCapability("platformName", "Windows");
            appOptions.AddAdditionalCapability("app", "Root");
            appOptions.AddAdditionalCapability("deviceName", "WindowsPC");
            WindowsDriver <WindowsElement> Session = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appOptions);

            WindowsElement applicationWindow = null;
            var openWindows = Session.FindElementsByClassName("Window");
            foreach (var window in openWindows)
            {
                Console.WriteLine(window.GetAttribute("Name"));
                if (window.GetAttribute("Name").StartsWith("SoapUI 5.5.0"))
                {
                    applicationWindow = window;
                    break;
                }
            }

            // Attaching to existing Application Window
            var topLevelWindowHandle = applicationWindow.GetAttribute("NativeWindowHandle");
            Console.WriteLine("Top level window handle " + topLevelWindowHandle);
            topLevelWindowHandle = int.Parse(topLevelWindowHandle).ToString("X");
            Console.WriteLine("Top level window handle after parse" + topLevelWindowHandle);

            AppiumOptions appOptions1 = new AppiumOptions();
            appOptions1.AddAdditionalCapability("deviceName", "WindowsPC");
            appOptions1.AddAdditionalCapability("appTopLevelWindow", topLevelWindowHandle);
            Session = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appOptions1);
            Console.WriteLine(Session.Title);
            Session.Close();
        }
    }
}
