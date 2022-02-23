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
                    par = context.Request.Query["klap_par"].ToString().Split(';');
                }
                String[] ram_par = null;
                if (context.Request.Query.ContainsKey("ram_par"))
                {
                    ram_par = context.Request.Query["ram_par"].ToString().Split(';');
                }           
                String[] priv_par = null;
                if (context.Request.Query.ContainsKey("priv_par"))
                {
                    priv_par = context.Request.Query["priv_par"].ToString().Split(';');
                }
                String[] backParts = null;
                if (context.Request.Query.ContainsKey("backParts"))
                {
                    backParts = context.Request.Query["backParts"].ToString().Split(';');
                }
                String[] frontParts = null;
                if (context.Request.Query.ContainsKey("frontParts"))
                {
                    frontParts = context.Request.Query["frontParts"].ToString().Split(';');
                }
                String[] hand = null;
                if (context.Request.Query.ContainsKey("hand"))
                {
                    hand = context.Request.Query["hand"].ToString().Split(';');
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
                    AcadBlockReference block = doc.ModelSpace.InsertBlock(new double[] { 100, 100, 0 }, path + @"DB\data.dwg", 1, 1, 1, 0);
                    AcadBlockReference block2 = doc.ModelSpace.InsertBlock(new double[] { 0, 0, 0 }, "Рамка", 1, 1, 1, 0);
                    object[] atrs2 = (object[])block2.GetAttributes();
                    await context.Response.WriteAsync(atrs2.Length.ToString() + "<Br>");
                    for (int i = 0; i < atrs2.Length; i++)
                    {
                        AcadAttributeReference atr;
                        atr = (AcadAttributeReference)atrs2[i];
                        atr.TextString = ram_par[i];
                        await context.Response.WriteAsync(ram_par[i] + "<Br>");
                    }
                    AcadBlockReference block6 = doc.ModelSpace.InsertBlock(new double[] { 150, 150, 0 }, "Привод_сбоку", 1, 1, 1, 0);
                    if (hand != null)
                    {
                        AcadBlockReference blockFor = doc.ModelSpace.InsertBlock(new double[] { 50, 150, 0 }, "part5", 1, 1, 1, 0);
                        object[] atrsFor = (object[])blockFor.GetAttributes();
                        await context.Response.WriteAsync(atrsFor.Length.ToString() + "<Br>");
                        for (int i = 0; i < atrsFor.Length; i++)
                        {
                            AcadAttributeReference atr;
                            atr = (AcadAttributeReference)atrsFor[i];
                            atr.TextString = priv_par[i + 2];
                            await context.Response.WriteAsync(priv_par[i+2] + "<Br>");
                        }
                    }
                    if(backParts != null)
                        foreach (string i in backParts)
                        {
                            AcadBlockReference blockFor = doc.ModelSpace.InsertBlock(new double[] { 50, 150, 0 }, "part" + i, 1, 1, 1, 0);
                        }
                    AcadBlockReference block1 = doc.ModelSpace.InsertBlock(new double[] { 50, 150, 0 }, name, 1, 1, 1, 0);
                    object[] atrs1 = (object[])block1.GetAttributes();
                    for (int i = 0; i < atrs1.Length; i++)
                    {
                        AcadAttributeReference atr;
                        atr = (AcadAttributeReference)atrs1[i];
                        atr.TextString = par[i];
                    }
                    AcadBlockReference block3 = doc.ModelSpace.InsertBlock(new double[] { 150 , 150 , 0 }, "Клапан_слева", 1, 1, 1, 0);
                    AcadBlockReference block4 = doc.ModelSpace.InsertBlock(new double[] { 50, 150, 0 }, "Привод", 1, 1, 1, 0);
                    object[] atrs4 = (object[])block4.GetAttributes();
                    await context.Response.WriteAsync(atrs4.Length.ToString() + "<Br>");
                    for (int i = 0; i < atrs4.Length; i++)
                    {
                        AcadAttributeReference atr;
                        atr = (AcadAttributeReference)atrs4[i];
                        atr.TextString = priv_par[i];
                        await context.Response.WriteAsync(priv_par[i] + "<Br>");
                    }
                    if (frontParts != null)
                        foreach (string i in frontParts)
                        {
                            AcadBlockReference blockFor = doc.ModelSpace.InsertBlock(new double[] { 50, 150, 0 }, "part" + i, 1, 1, 1, 0);
                        }
                    var acPlotCfg = doc.PlotConfigurations;
                    acPlotCfg.Add("PDF", true); // If second parameter is not true, exception is caused by acDoc.ActiveLayout.CopyFrom(PlotConfig);
                    AcadPlotConfiguration PlotConfig = acPlotCfg.Item("PDF");
                    
                    PlotConfig.ConfigName = "AutoCAD PDF (General Documentation).pc3";
                    PlotConfig.CanonicalMediaName = "ISO_full_bleed_A4_(210.00_x_297.00_MM)";
                    PlotConfig.PlotHidden = false;
                    PlotConfig.StandardScale = AcPlotScale.ac1_1;
                    PlotConfig.PlotType = AcPlotType.acLimits;
                    PlotConfig.PaperUnits = AcPlotPaperUnits.acMillimeters;
                    PlotConfig.PlotRotation = AcPlotRotation.ac0degrees;
                    PlotConfig.CenterPlot = true;
                    PlotConfig.PlotOrigin = new double[2] { 12.5 , 5 };
                    PlotConfig.PlotWithLineweights = false;
                    PlotConfig.PlotWithPlotStyles = false;
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
