using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;
using Autodesk.AutoCAD;
using System.IO;

namespace test2
{
    public class Startup
    {
        // This method gets called by the runtime. Use this method to add services to the container.
        // For more information on how to configure your application, visit https://go.microsoft.com/fwlink/?LinkID=398940
        static AcadApplication acad = null;
        static string path = Directory.GetParent(Directory.GetParent(Directory.GetCurrentDirectory()).FullName).FullName + @"\";
        public void ConfigureServices(IServiceCollection services)
        {
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app)
        {
            app.Map("/index", Index);
            app.Map("/about", About);

            app.Run(async (context) =>
            {
                await context.Response.WriteAsync("Page Not Found");
            });
        }

        private static void Index(IApplicationBuilder app)
        {
            
            app.Use(async (context, next) =>
            {
                 String name = null;
                if (context.Request.Query.ContainsKey("klap"))
                {
                    name = context.Request.Query["klap"];
                }
                String[] par = null;
                if (context.Request.Query.ContainsKey("klap_par"))
                {
                    par = context.Request.Query["klap_par"].ToString().Split(',');
                }
                await context.Response.WriteAsync("<p>Hello world!</p>");
                
                while (acad == null)
                    try
                    {
                        acad = new AcadApplication();
                    }
                    catch (Exception e)
                    {
                        await context.Response.WriteAsync(e.ToString() + "<Br>");
                    }
                await context.Response.WriteAsync("acad open" + "<Br>");
                AcadDocument doc = null;
                while (doc == null)
                    try
                    {
                        doc = acad.Documents.Add();
                    }
                    catch (Exception e)
                    {
                        await context.Response.WriteAsync(e.ToString() + "<Br>");
                    }
                await context.Response.WriteAsync("doc create" + "<Br>");
                try
                {
                    doc.ActiveSpace = AcActiveSpace.acModelSpace;
                    Double[] t1 = { 100, 100, 0 };
                    AcadBlockReference block = doc.ModelSpace.InsertBlock(t1, path + @"DB\data.dwg", 1, 1, 1, 0);
                    //block.Delete();
                    AcadBlockReference block1 = doc.ModelSpace.InsertBlock(t1, name, 1, 1, 1, 0);
                    //var att = block1.GetAttributes();
                    object[] a = (object[])block1.GetAttributes();
                    for (int i = 0; i < a.Length; i++)
                    {
                        AcadAttributeReference atr;
                        atr = (AcadAttributeReference)a[i];
                        atr.TextString = par[i];
                    }
                    //doc.SaveAs(@"..\..\..\..\testApi.dwg");

                    var acPlotCfg = doc.PlotConfigurations;
                    acPlotCfg.Add("PDF", true); // If second parameter is not true, exception is caused by acDoc.ActiveLayout.CopyFrom(PlotConfig);
                    AcadPlotConfiguration PlotConfig = acPlotCfg.Item("PDF");
                    
                    PlotConfig.ConfigName = "DWG To PDF.pc3";
                    PlotConfig.CanonicalMediaName = "ISO_A4_(297.00_x_210.00_MM)";
                    PlotConfig.PlotHidden = false;
                    PlotConfig.StandardScale = AcPlotScale.acScaleToFit;
                    PlotConfig.PlotType = AcPlotType.acLimits;
                    PlotConfig.PaperUnits = AcPlotPaperUnits.acMillimeters;
                    PlotConfig.PlotRotation = AcPlotRotation.ac90degrees;
                    PlotConfig.CenterPlot = true;
                    PlotConfig.PlotOrigin = new double[2] { 0, 0 };
                    PlotConfig.PlotWithLineweights = false;
                    PlotConfig.PlotWithPlotStyles = false;
                    //PlotConfig.StyleSheet = "acad.ctb";
                    PlotConfig.RefreshPlotDeviceInfo();
                    
                    doc.ActiveLayout.CopyFrom(PlotConfig); // Need to have this or resulting PDF does not seem to apply the PlotConfig.
                    doc.SetVariable("BACKGROUNDPLOT", 1);
                    doc.Plot.QuietErrorMode = true;
                    doc.Plot.NumberOfCopies = 1;
                    doc.Plot.PlotToFile(path + @"testApi.pdf", PlotConfig.ConfigName);
                    


                }
                catch (Exception e)
                {
                    await context.Response.WriteAsync(e.ToString()+ "<Br>");
                }
                finally
                {
                    doc.SaveAs(path + @"testApi.dwg");
                    doc.Close();
                    await context.Response.WriteAsync("<p>gg</p>");
                }
                

                await next.Invoke();
            });
            

            app.Run(async context =>
            {
                await context.Response.WriteAsync("<p>fin</p>");

            });
        }
        private static void About(IApplicationBuilder app)
        {
            app.Run(async context =>
            {
                await context.Response.WriteAsync(path + @"DB\data.dwg");
            });
        }
    }
}
