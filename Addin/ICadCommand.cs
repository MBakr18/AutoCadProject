﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddIn
{
  public abstract class ICadCommand
  {
    /// <summary>
    /// Execute a function in cad application
    /// </summary>
    public abstract void Execute();
  }
}
