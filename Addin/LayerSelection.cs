using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Addin
{
  public class LayerSelection :ICadCommand
  {
    public override void Execute()
    {
      SelectFromScreen();
    }

   
  }
}
