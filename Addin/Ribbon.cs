using CadProject;
using Autodesk.AutoCAD.GraphicsInterface;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using Exception = System.Exception;

namespace AddIn
{
  public class Ribbon
  {
    public const string RibbonTitle = "MB Tools";
    public const string RibbonId = "10 10";

    [CommandMethod("RibbonCreator")]
    public void CreateRibbon()
    {
      RibbonControl ribbon = ComponentManager.Ribbon;
      if (ribbon != null)
      {
        RibbonTab rtab = ribbon.FindTab(RibbonId);
        if (rtab != null)
        {
          ribbon.Tabs.Remove(rtab);
        }

        rtab = new RibbonTab();
        rtab.Title = RibbonTitle;
        rtab.Id = RibbonId;
        ribbon.Tabs.Add(rtab);
        AddContentToTab(rtab);
      }
    }

    private void AddContentToTab(RibbonTab rtab)
    {
      rtab.Panels.Add(AddPanelOne());
    }

    private static RibbonPanel AddPanelOne()
    {
      RibbonPanelSource rps = new RibbonPanelSource();
      rps.Title = "Cols Creation";
      RibbonPanel rp = new RibbonPanel();
      rp.Source = rps;
      RibbonButton rci = new RibbonButton();
      rci.Name = "MB Addin";
      rps.DialogLauncher = rci;


      //create button1
      var addinAssembly = typeof(Ribbon).Assembly;
      RibbonButton gridCreation = new RibbonButton
      {
        Orientation = Orientation.Vertical,
        AllowInStatusBar = true,
        Size = RibbonItemSize.Large,
        Name = "gridCreationBtn",
        ShowText = true,
        Text = "Import Grid System\nFrom Excel",
        Description = "Import grids From Excel File",
        CommandHandler = new RelayCommand(new ImportAndDrawGrid().Execute),
        // Image = GetEmbeddedPng(typeof(Ribbon).Assembly, "AddIn.Resources.ITI.png")
      };

      //create button2
      RibbonButton colCreation = new RibbonButton
      {
          Orientation = Orientation.Vertical,
          AllowInStatusBar = true,
          Size = RibbonItemSize.Large,
          Name = "colCreationBtn",
          ShowText = true,
          Text = "Draw Cols\nAt Grid Intersection",
          Description = "Add Cols at grid interstion in the model",
          CommandHandler = new RelayCommand(new DrawColsAtGrids().Execute),
          // Image = GetEmbeddedPng(typeof(Ribbon).Assembly, "AddIn.Resources.ITI.png")
      };

      rps.Items.Add(gridCreation);
      rps.Items.Add(colCreation);

      return rp;
    }

    public static System.Windows.Media.ImageSource GetEmbeddedPng(System.Reflection.Assembly app, string imageName)
    {
        var file = app.GetManifestResourceStream(imageName);
        BitmapDecoder source = PngBitmapDecoder.Create(file, BitmapCreateOptions.None, BitmapCacheOption.None);
        return source.Frames[0];
    }
  }
}


