using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoboSharp.BackupApp {
  public static class ByteRate {
    const int ARRAYLENGTH = 50;
    static double[] bytesArray = new double[ARRAYLENGTH];
    static double[] msArray = new double[ARRAYLENGTH];
    static int ArrayCount;
    static int ArrayIn;

    public static void Reset() {
      ArrayCount = 0;
      ArrayIn = 0;
    }

    public static double Add(double bytes, 
      double ms) {
      double bytes1;
      double ms1;
      if (ArrayCount < ARRAYLENGTH) {
        ++ArrayCount;
        bytes1 = bytes - bytesArray[0];
        ms1 = ms - msArray[0];
      }
      else {
        bytes1 = bytes - bytesArray[ArrayIn];
        ms1 = ms - msArray[ArrayIn];
      }
      bytesArray[ArrayIn] = bytes;
      msArray[ArrayIn] = ms;
      ArrayIn = ArrayIn + 1 < ARRAYLENGTH ? ++ArrayIn : 0;
      if (ms1 < 1000) {
        return 0.0;
      }
      else {
        double mbs = bytes1 * 1000 / 1024.0 / 1024.0 / ms1;
        return mbs;
      }
    }
  }
}
