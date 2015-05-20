﻿using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioCSharpSamples
{

    public static partial class Samples
    {
        public static void Page_SetResults(IVisio.Document doc)
        {
            var page = Util.CreateStandardPage(doc, "PSR");
            var shape = Util.CreateStandardShape(page);

            // CREATE REQUEST
            var request = new[]
                              {
                                  new
                                      {
                                          ID = (short) shape.ID16,
                                          Section = (short) IVisio.VisSectionIndices.visSectionObject,
                                          Row = (short) IVisio.VisRowIndices.visRowXFormOut,
                                          Cell = (short) IVisio.VisCellIndices.visXFormWidth,
                                          UnitCode = (short) IVisio.VisUnitCodes.visNoCast,
                                          Result = (double) 8.0
                                      },
                                  new
                                      {
                                          ID = (short) shape.ID16,
                                          Section = (short) IVisio.VisSectionIndices.visSectionObject,
                                          Row = (short) IVisio.VisRowIndices.visRowXFormOut,
                                          Cell = (short) IVisio.VisCellIndices.visXFormHeight,
                                          UnitCode = (short) IVisio.VisUnitCodes.visNoCast,
                                          Result = (double) 1.3
                                      }
                              };

            // MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
            var SID_SRCStream = new short[request.Length*4];
            var results_objects = new object[request.Length];
            var unitcodes = new object[request.Length];
            for (int i = 0; i < request.Length; i++)
            {
                SID_SRCStream[(i*4) + 0] = request[i].ID;
                SID_SRCStream[(i*4) + 1] = request[i].Section;
                SID_SRCStream[(i*4) + 2] = request[i].Row;
                SID_SRCStream[(i*4) + 3] = request[i].Cell;
                results_objects[i] = request[i].Result;
                unitcodes[i] = request[i].UnitCode;
            }

            // EXECUTE THE REQUEST
            short flags = 0;
            int count = page.SetResults(SID_SRCStream, unitcodes, results_objects, flags);

            // DISPLAY THE INFORMATION
            shape.Text = "SetResults";
        }
    }

}