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
using System.Text.Json;

namespace test2
{
    public class Startup
    {
        // This method gets called by the runtime. Use this method to add services to the container.
        // For more information on how to configure your application, visit https://go.microsoft.com/fwlink/?LinkID=398940

        // ������ �� ��� X.
        private const int _xScale = 1;
        // ������ �� ��� Y.
        private const int _yScale = 1;
        // ������ �� ��� Z.
        private const int _zScale = 1;
        // ������� �����.
        private const int _rotate = 0;
        static AcadApplication acadApp = null;
        static string path = Directory.GetParent(Directory.GetParent(Directory.GetCurrentDirectory()).FullName).FullName + @"\";
        public void ConfigureServices(IServiceCollection services)
        {
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app)
        {
            app.Map("/index", Index);
            app.Map("/about", About);
            app.Map("/jsonTest", JsonTest);

            app.Run(async (context) =>
            {
                await context.Response.WriteAsync("Page Not Found");
            });
        }
        class kord
        {
            public int X { get; set; }
            public int Y { get; set; }
            public kord(int x, int y)
            {
                X = x;
                Y = y;
            }
        }
        class ram
        {
            public string Name { get;}  
            public kord ClapXY { get; set; }
            public ram(string name, kord clapXY)
            {
                Name = name;
                ClapXY = clapXY;
            }

        }
        private static void JsonTest(IApplicationBuilder app)
        {
            app.Use(async (context, next) =>
            {
                ram test = new ram("test", new kord (12,12));
                string json = JsonSerializer.Serialize(test);
                await context.Response.WriteAsync(json);
            });

        }
            private static void Index(IApplicationBuilder app)
        {
            app.Use(async (context, next) =>
            {
                // �������� ��� ����
                // � ������ ������� ���������� ������ � ������ �������� �������� � ��������� �������, ������� ���� ������.

                // �������� �������.
                String name = null;
                if (context.Request.Query.ContainsKey("klap"))
                {
                    name = context.Request.Query["klap"];
                }

                // ��������� �������.
                String[] valveParameters = null;
                if (context.Request.Query.ContainsKey("klap_par"))
                {
                    valveParameters = context.Request.Query["klap_par"].ToString().Split(';');
                }

                // ����������� 
                String[] frameParameters = null;
                if (context.Request.Query.ContainsKey("ram_par"))
                {
                    frameParameters = context.Request.Query["frameParameters"].ToString().Split(';');
                }     
                
                // ��������� �������
                String[] actuatorParameters = null;
                if (context.Request.Query.ContainsKey("priv_par"))
                {
                    actuatorParameters = context.Request.Query["actuatorParameters"].ToString().Split(';');
                }

                // ����� �� ������ ����.
                String[] backParts = null;
                if (context.Request.Query.ContainsKey("backParts"))
                {
                    backParts = context.Request.Query["backParts"].ToString().Split(';');
                }

                // ����� �� �������� ����.
                String[] frontParts = null;
                if (context.Request.Query.ContainsKey("frontParts"))
                {
                    frontParts = context.Request.Query["frontParts"].ToString().Split(';');
                }

                // ���������� ������� ���������� ������� ������� ������� �� �������
                String[] handActuator = null;
                if (context.Request.Query.ContainsKey("hand"))
                {
                    handActuator = context.Request.Query["handActuator"].ToString().Split(';');
                }

                // �������� ����� �������� �������
                String filename = null;
                if (context.Request.Query.ContainsKey("filename"))
                {
                    filename = context.Request.Query["filename"];
                }
                
                // ��������� �������
                while (acadApp == null)
                    try
                    {
                        acadApp = new AcadApplication();
                    }
                    catch (Exception e)
                    {
                        await context.Response.WriteAsync(e.ToString() + "<Br>");
                    }
                await context.Response.WriteAsync("acad open" + "<Br>");

                // ������ �������� � ��������
                AcadDocument acadDoc = null;
                while (acadDoc == null)
                    try
                    {
                        acadDoc = acadApp.Documents.Add();
                    }
                    catch (Exception e)
                    {
                        await context.Response.WriteAsync(e.ToString() + "<Br>");
                    }
                
                await context.Response.WriteAsync("doc create" + "<Br>");
                try
                {
                    acadDoc.ActiveSpace = AcActiveSpace.acModelSpace;

                    // ��� �� ��������� ���� ( ��������� ����� ���� ��������� ������� ) � ������� � ��� ����� �� � ��������� � ���� ������
                    AcadBlockReference acadBlock = acadDoc.ModelSpace.InsertBlock(new double[] { 0, 0, 0 }, path + @"DB\data.dwg",
                        _xScale, _yScale, _zScale, _rotate);
                    // ��������� ���� ����� �� ����������� ����������
                    AcadBlockReference acadBlockFrame = acadDoc.ModelSpace.InsertBlock(new double[] { 0, 0, 0 }, "�����",
                        _xScale, _yScale, _zScale, _rotate);
                    object[] acadBlockFrameAttributes = (object[])acadBlockFrame.GetAttributes();
                    await context.Response.WriteAsync(acadBlockFrameAttributes.Length.ToString() + "<Br>");
                    for (int i = 0; i < acadBlockFrameAttributes.Length; i++)
                    {
                        AcadAttributeReference acadBlockFrameAttribute = (AcadAttributeReference)acadBlockFrameAttributes[i];
                        acadBlockFrameAttribute.TextString = frameParameters[i];
                        await context.Response.WriteAsync(frameParameters[i] + "<Br>");
                    }

                    // ���������� �����, ������� ���������� ������ �����.
                    AcadBlockReference acadBlockActuatorSideView = acadDoc.ModelSpace.InsertBlock(new double[] { 150, 150, 0 }, "������_�����",
                        _xScale, _yScale, _zScale, _rotate);

                    // ���������� �����, ������� ���������� ������ ������ �� �������.
                    if (handActuator != null)
                    {
                        AcadBlockReference acadBlockHandActuator = acadDoc.ModelSpace.InsertBlock(new double[] { 50, 150, 0 }, "part5",
                            _xScale, _yScale, _zScale, _rotate);
                        object[] acadBlockHandActuatorAttributes = (object[])acadBlockHandActuator.GetAttributes();
                        await context.Response.WriteAsync(acadBlockHandActuatorAttributes.Length.ToString() + "<Br>");
                        for (int i = 0; i < acadBlockHandActuatorAttributes.Length; i++)
                        {
                            AcadAttributeReference acadBlockHandActuatorAttribute;
                            acadBlockHandActuatorAttribute = (AcadAttributeReference)acadBlockHandActuatorAttributes[i];
                            // ��������� ������� ������� ����� �������� � ���������� ������� � ����� ������� ( ������� ��� ������������ ����������� i + 2 )
                            acadBlockHandActuatorAttribute.TextString = actuatorParameters[i + 2];
                            await context.Response.WriteAsync(actuatorParameters[i+2] + "<Br>");
                        }
                    }

                    // ���������� �����, ������� ���������� �������� ������������ �� ������ ���� �������.
                    if (backParts != null)
                        foreach (string i in backParts)
                        {
                            AcadBlockReference acadBlockBackPart = acadDoc.ModelSpace.InsertBlock(new double[] { 50, 150, 0 }, "part" + i,
                                _xScale, _yScale, _zScale, _rotate);
                        }

                    // ���������� �����, ������� ���������� ������ �� �������.
                    AcadBlockReference acadBlockValve = acadDoc.ModelSpace.InsertBlock(new double[] { 50, 150, 0 }, name,
                        _xScale, _yScale, _zScale, _rotate); ;
                    object[] acadBlockValveAttributes = (object[])acadBlockValve.GetAttributes();
                    for (int i = 0; i < acadBlockValveAttributes.Length; i++)
                    {
                        AcadAttributeReference acadBlockValveAttribute;
                        acadBlockValveAttribute = (AcadAttributeReference)acadBlockValveAttributes[i];
                        acadBlockValveAttribute.TextString = valveParameters[i];
                    }

                    // ���������� �����, ������� ���������� ������ ( ��� ����� ) �� �������.
                    AcadBlockReference acadBlockValveSideView = acadDoc.ModelSpace.InsertBlock(new double[] { 150 , 150 , 0 }, "������_�����",
                        _xScale, _yScale, _zScale, _rotate); ;

                    // ���������� �����, ������� ���������� ������ �� �������.
                    AcadBlockReference acadBlockAcruator = acadDoc.ModelSpace.InsertBlock(new double[] { 50, 150, 0 }, "������",
                    _xScale, _yScale, _zScale, _rotate); ;
                    object[] acadBlockAcruatorAttributes = (object[])acadBlockAcruator.GetAttributes();
                    for (int i = 0; i < acadBlockAcruatorAttributes.Length; i++)
                    {
                        AcadAttributeReference acadBlockAcruatorAttribute;
                        acadBlockAcruatorAttribute = (AcadAttributeReference)acadBlockAcruatorAttributes[i];
                        acadBlockAcruatorAttribute.TextString = actuatorParameters[i];
                    }

                    // ���������� �����, ������� ���������� �������� ������������ �� �������� ���� �������.
                    if (frontParts != null)
                        foreach (string i in frontParts)
                        {
                            AcadBlockReference acadBlockFrontParts = acadDoc.ModelSpace.InsertBlock(new double[] { 50, 150, 0 }, "part" + i,
                                _xScale, _yScale, _zScale, _rotate); ;
                        }

                    // ��������� ���������� ��� ������ ������� � ������� PDF.
                    var acPlotCfg = acadDoc.PlotConfigurations;
                    acPlotCfg.Add("PDF", true);
                    AcadPlotConfiguration plotConfig = acPlotCfg.Item("PDF");
                    
                    plotConfig.ConfigName = "AutoCAD PDF (General Documentation).pc3";
                    plotConfig.CanonicalMediaName = "ISO_full_bleed_A4_(210.00_x_297.00_MM)";
                    plotConfig.PlotHidden = false;
                    plotConfig.StandardScale = AcPlotScale.ac1_1;
                    plotConfig.PlotType = AcPlotType.acLimits;
                    plotConfig.PaperUnits = AcPlotPaperUnits.acMillimeters;
                    plotConfig.PlotRotation = AcPlotRotation.ac0degrees;
                    plotConfig.CenterPlot = true;
                    plotConfig.PlotOrigin = new double[2] { 12.5 , 5 };
                    plotConfig.PlotWithLineweights = false;
                    plotConfig.PlotWithPlotStyles = false;
                    plotConfig.RefreshPlotDeviceInfo();
                    acadDoc.ActiveLayout.CopyFrom(plotConfig);
                    acadDoc.SetVariable("BACKGROUNDPLOT", 1);
                    acadDoc.Plot.QuietErrorMode = true;
                    acadDoc.Plot.NumberOfCopies = 1;
                    acadDoc.Plot.PlotToFile(path + filename + ".pdf", plotConfig.ConfigName);
                }
                catch (Exception e)
                {
                    await context.Response.WriteAsync(e.ToString()+ "<Br>");
                }
                finally
                {
                    acadDoc.SaveAs(path + filename + ".dwg");
                    acadDoc.Close();
                }
                await next.Invoke();
            });
        }
        private static void About(IApplicationBuilder app)
        {
            app.Run(async context =>
            {
                await context.Response.WriteAsync(path + @"DB\data.dwg").ConfigureAwait(false);
            });
        }

        ~Startup()
        {
            Console.WriteLine("qwqe");
            if (acadApp != null)
                acadApp.Quit();
            Console.WriteLine("131");

        }
    }
}
