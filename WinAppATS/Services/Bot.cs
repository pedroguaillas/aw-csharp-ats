using System;
using RestSharp;
using OpenQA.Selenium;
using System.Threading;
using System.Windows.Forms;
using OpenQA.Selenium.Chrome;
using System.Collections.Generic;

namespace WinAppATS.Services
{
    class Bot : Importa
    {
        string apiKey = "3a07e3f65175000d8e2baa64ff999700";
        string[] errors = { "No existen datos para los parámetros  ingresados ", "Captcha incorrecta" };
        public void DescargarRecibidos(string info, string pass, string ano, string mes, string tc, int ta)
        {
            string folder = selectFolder();

            if (folder == null)
            {
                return;
            }

            try
            {
                var chromeDriverService = ChromeDriverService.CreateDefaultService();
                chromeDriverService.HideCommandPromptWindow = true;
                var optionChrome = new ChromeOptions();

                optionChrome.AddArgument("--headless=new");
                optionChrome.AddArgument("window-size=0x0");

                optionChrome.AddUserProfilePreference("profile.default_content_setting_values.automatic_downloads", 1);
                optionChrome.AddUserProfilePreference("safebrowsing.enabled", true);
                optionChrome.AddUserProfilePreference("download.default_directory", folder);

                using (ChromeDriver driver = new ChromeDriver(chromeDriverService, optionChrome))
                {
                    // Esperara hasta 3 min que se cargue la pagina
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(3);
                    driver.Navigate().GoToUrl("https://srienlinea.sri.gob.ec/auth/realms/Internet/protocol/openid-connect/auth?response_type=code&client_id=app-sri-claves-angular");

                    // Establcer credenciales
                    var txtUsuario = driver.FindElement(By.Id("usuario"));
                    txtUsuario.SendKeys(info.Substring(0, 13));
                    var txtPassword = driver.FindElement(By.Id("password"));
                    txtPassword.SendKeys(pass);

                    // Logearse
                    var btnLogin = driver.FindElement(By.Id("kc-login"));
                    btnLogin.Click();

                    //Espera que se termine de logear
                    Thread.Sleep(3000);

                    // IR a la pagina de consulta de comprobantes RECIBIDAS
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(3);
                    driver.Navigate().GoToUrl("https://srienlinea.sri.gob.ec/tuportal-internet/accederAplicacion.jspa?redireccion=57&idGrupo=55");

                    // Selecionamos el año
                    var selAno = driver.FindElement(By.Id("frmPrincipal:ano"));
                    IReadOnlyCollection<IWebElement> anos = selAno.FindElements(By.TagName("option"));
                    foreach (IWebElement option in anos)
                    {
                        if (option.GetAttribute("value").Equals(ano)) // Año del 2023
                            option.Click();
                    }
                    // Selecionamos el mes
                    var selMes = driver.FindElement(By.Id("frmPrincipal:mes"));
                    IReadOnlyCollection<IWebElement> meses = selMes.FindElements(By.TagName("option"));
                    foreach (IWebElement m in meses)
                    {
                        if (m.GetAttribute("value").Equals(mes)) // Mes de enero
                            m.Click();
                    }
                    // Ponemos todos los dias
                    var selDia = driver.FindElement(By.Id("frmPrincipal:dia"));
                    IReadOnlyCollection<IWebElement> dias = selDia.FindElements(By.TagName("option"));
                    foreach (IWebElement dia in dias)
                    {
                        if (dia.GetAttribute("value").Equals("0")) // Todos los dias
                            dia.Click();
                    }

                    // Seleccionamos el tipo de comprobante
                    var selComprobante = driver.FindElement(By.Id("frmPrincipal:cmbTipoComprobante"));
                    IReadOnlyCollection<IWebElement> comprobantes = selComprobante.FindElements(By.TagName("option"));
                    foreach (IWebElement option in comprobantes)
                    {
                        if (option.GetAttribute("value").Equals(tc))  // 1 Factura
                            option.Click();
                    }

                    //Pasar Captcha
                    var data_sitekey = driver.ExecuteScript("return document.getElementsByClassName('g-recaptcha')[0].getAttribute('data-sitekey')");
                    string pageurl = "https://srienlinea.sri.gob.ec";

                    //REQUEST IN
                    var client = new RestClient("https://2captcha.com");
                    var requestIn = new RestRequest("in.php", Method.Post);
                    requestIn.AddBody(new
                    {
                        key = apiKey,
                        method = "userrecaptcha",
                        googlekey = data_sitekey.ToString(),
                        pageurl
                    });
                    var resIn = client.ExecutePost(requestIn);
                    string captcha_id = resIn.Content.ToString().Split('|')[1];

                    //REQUEST RES
                    string captcha_res = "";

                    do
                    {
                        var clientRes = new RestClient("http://2captcha.com");
                        var requestRes = new RestRequest("res.php", Method.Get);

                        requestRes.AddParameter("key", apiKey);
                        requestRes.AddParameter("action", "get");
                        requestRes.AddParameter("id", captcha_id);

                        var resRes = clientRes.Execute(requestRes);

                        if (resRes.Content != null)
                        {
                            string res = resRes.Content.ToString();
                            if (res.Contains("|"))
                            {
                                captcha_res = res.Split('|')[1];
                            }
                            else
                            {
                                Thread.Sleep(5000);
                            }
                        }
                        else
                        {
                            MessageBox.Show("El dato es nulo");
                        }

                    } while (captcha_res.Length == 0);

                    // Consultar comprobantes
                    driver.ExecuteScript(String.Format("document.getElementById('g-recaptcha-response').value='{0}'", captcha_res));
                    driver.ExecuteScript(String.Format("rcBuscar('{0}')", captcha_res));

                    // Consultar comprobantes sin CAPTCHA
                    //driver.ExecuteScript("grecaptcha.execute()");

                    //ESPERAR 5 min por que aqui se demora un poco mas con respecto a los demas ESPERAS.
                    Thread.Sleep(5000);

                    // Verificar SI EXISTE error
                    var haserror = (bool)driver.ExecuteScript("return document.getElementById('formMessages:messages').hasChildNodes()");
                    if (haserror)
                    {
                        var warning = driver.ExecuteScript("return document.getElementsByClassName('ui-messages-warn-summary')[0].textContent");

                        //Sino existe registros finalizar
                        if (warning.Equals(errors[0]))
                        {
                            driver.Close();
                            MessageBox.Show("Se finalizo la descarga");
                            return;
                        }

                        MessageBox.Show(warning.ToString());
                    }

                    // Seleccionamos la tabla
                    var tableComprobantes = driver.FindElement(By.Id("frmPrincipal:tablaCompRecibidos_data"));
                    var rows = tableComprobantes.FindElements(By.TagName("tr")).Count;

                    string page_current = driver.FindElement(By.ClassName("ui-paginator-current")).Text;
                    page_current = page_current.Replace("(", "");
                    page_current = page_current.Replace(")", "");
                    string[] pags = page_current.Split(' ');

                    const int size_page = 300;
                    int last_page = int.Parse(pags[2].ToString());

                    //sim tiene mas de 1 pagina entonces cargara hasta 300 comprobantes por pagina
                    if (last_page > 1)
                    {
                        driver.ExecuteScript(string.Format("let max_results = {0}; {1} {2} {3} {4} {5} {6} {7} {8}",
                              size_page,
                              "var select = document.getElementById('frmPrincipal:tablaCompRecibidos_paginator_bottom').childNodes[0];",
                              "var option = document.createElement('option');",
                              "option.text = max_results;",
                              "option.value = max_results;",
                              "select.add(option);",
                              "select.selectedIndex = select.options.length - 1;",
                              "select.dispatchEvent(new Event('change'));",
                              "select.dispatchEvent(new Event('change'));")
                              );

                        Thread.Sleep(3000);

                        //UNA VEZ QUE SE RECARGA SE DEBE VOLVER A CALCULAR LA ULTIMA PÁGINA
                        page_current = driver.FindElement(By.ClassName("ui-paginator-current")).Text;
                        page_current = page_current.Replace("(", "");
                        page_current = page_current.Replace(")", "");
                        pags = page_current.Split(' ');
                        last_page = int.Parse(pags[2].ToString());
                    }

                    // Contador de numero de comprobantes
                    int i = 0;

                    //Recorre la cantidad de paginas
                    for (int current = 1; current <= last_page; current++)
                    {
                        tableComprobantes = driver.FindElement(By.Id("frmPrincipal:tablaCompRecibidos_data"));
                        rows = tableComprobantes.FindElements(By.TagName("tr")).Count;

                        if (rows > 0)
                        {
                            for (int ir = 0; ir < rows; ir++)
                            {
                                if (ta == 0)
                                {
                                    driver.ExecuteScript(string.Format("document.getElementById('frmPrincipal:tablaCompRecibidos:{0}:lnkXml').click();", i++));
                                }
                                else
                                {
                                    driver.ExecuteScript(string.Format("document.getElementById('frmPrincipal:tablaCompRecibidos:{0}:lnkPdf').click();", i++));
                                }
                                Thread.Sleep(500);
                            }

                            if (current < last_page)
                            {
                                driver.FindElement(By.ClassName("ui-paginator-next")).Click();
                                //driver.ExecuteScript("document.getElementsByClassName('ui-paginator-next')[0].click();");
                                Thread.Sleep(7500);
                            }
                        }
                    }

                    driver.Close();
                    MessageBox.Show("Se finalizo la descarga");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error " + ex.Message);
            }
        }

        public void DescargaEmitodos(string info, string pass, string fecha, string tc)
        {
            string folder = selectFolder();

            if (folder == null)
            {
                return;
            }

            try
            {
                var chromeDriverService = ChromeDriverService.CreateDefaultService();
                chromeDriverService.HideCommandPromptWindow = true;
                var optionChrome = new ChromeOptions();

                optionChrome.AddArgument("--headless=new");

                optionChrome.AddUserProfilePreference("profile.default_content_setting_values.automatic_downloads", 1);
                optionChrome.AddUserProfilePreference("safebrowsing.enabled", true);
                optionChrome.AddUserProfilePreference("download.default_directory", folder);

                using (ChromeDriver driver = new ChromeDriver(chromeDriverService, optionChrome))
                {
                    // Esperara hasta 3 min que se cargue la pagina
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(3);
                    driver.Navigate().GoToUrl("https://srienlinea.sri.gob.ec/auth/realms/Internet/protocol/openid-connect/auth?response_type=code&client_id=app-sri-claves-angular");

                    // Establecer credenciales
                    var txtUsuario = driver.FindElement(By.Id("usuario"));
                    txtUsuario.SendKeys(info.Substring(0, 13));
                    var txtPassword = driver.FindElement(By.Id("password"));
                    txtPassword.SendKeys(pass);

                    // Logearse
                    var btnLogin = driver.FindElement(By.Id("kc-login"));
                    btnLogin.Click();

                    //Espera que se termine de logear
                    Thread.Sleep(3000);

                    // IR a la pagina de FACTURACION ELECTRONICA Consultas
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(3);
                    driver.Navigate().GoToUrl("https://srienlinea.sri.gob.ec/tuportal-internet/accederAplicacion.jspa?redireccion=60&idGrupo=58");

                    // IR a la pagina de consulta de comprobantes EMITIDAS
                    driver.ExecuteScript("mojarra.jsfcljs(document.getElementById('consultaDocumentoForm'),{'consultaDocumentoForm:j_idt22':'consultaDocumentoForm:j_idt22'},'');");

                    //Esperar que se cargue el formulario de selección de comprobantes
                    Thread.Sleep(3000);

                    // IR a la pagina de consulta de comprobantes EMITIDAS
                    driver.ExecuteScript(String.Format("document.getElementById('frmPrincipal:calendarFechaDesde_input').value='{0}'", fecha));

                    // Seleccionamos el tipo de comprobante
                    var selComprobante = driver.FindElement(By.Id("frmPrincipal:cmbTipoComprobante"));
                    IReadOnlyCollection<IWebElement> comprobantes = selComprobante.FindElements(By.TagName("option"));
                    foreach (IWebElement option in comprobantes)
                    {
                        if (option.GetAttribute("value").Equals(tc))  // 1 Factura
                            option.Click();
                    }

                    //Pasar Captcha
                    var data_sitekey = driver.ExecuteScript("return document.getElementsByClassName('g-recaptcha')[0].getAttribute('data-sitekey')");
                    string pageurl = "https://srienlinea.sri.gob.ec";

                    //REQUEST IN
                    var client = new RestClient("https://2captcha.com");
                    var requestIn = new RestRequest("in.php", Method.Post);
                    requestIn.AddBody(new
                    {
                        key = apiKey,
                        method = "userrecaptcha",
                        googlekey = data_sitekey.ToString(),
                        pageurl
                    });
                    var resIn = client.ExecutePost(requestIn);
                    string captcha_id = resIn.Content.ToString().Split('|')[1];

                    //REQEUST RES
                    string captcha_res = "";

                    do
                    {
                        var clientRes = new RestClient("http://2captcha.com");
                        var requestRes = new RestRequest("res.php", Method.Get);

                        requestRes.AddParameter("key", apiKey);
                        requestRes.AddParameter("action", "get");
                        requestRes.AddParameter("id", captcha_id);

                        var resRes = clientRes.Execute(requestRes);

                        if (resRes.Content != null)
                        {
                            string res = resRes.Content.ToString();
                            if (res.Contains("|"))
                            {
                                captcha_res = res.Split('|')[1];
                            }
                            else
                            {
                                Thread.Sleep(5000);
                            }
                        }
                        else
                        {
                            MessageBox.Show("El dato es nulo");
                        }

                    } while (captcha_res.Length == 0);

                    // Consultar comprobantes
                    driver.ExecuteScript(String.Format("document.getElementById('g-recaptcha-response').value='{0}'", captcha_res));
                    driver.ExecuteScript(String.Format("rcBuscar('{0}')", captcha_res));

                    //ESPERAR
                    Thread.Sleep(3000);

                    // Verificar SI EXISTE error
                    var haserror = (bool)driver.ExecuteScript("return document.getElementById('formMessages:messages').hasChildNodes()");
                    if (haserror)
                    {
                        var warning = driver.ExecuteScript("return document.getElementsByClassName('ui-messages-warn-summary')[0].textContent");

                        //Sino existe registros finalizar
                        if (warning.Equals(errors[0]))
                        {
                            driver.Close();
                            MessageBox.Show("Se finalizo la descarga");
                            return;
                        }

                        MessageBox.Show(warning.ToString());
                    }

                    // Seleccionamos la tabla
                    var tableComprobantes = driver.FindElement(By.Id("frmPrincipal:tablaCompEmitidos_data"));
                    var rows = tableComprobantes.FindElements(By.TagName("tr")).Count;

                    string page_current = driver.FindElement(By.ClassName("ui-paginator-current")).Text;
                    page_current = page_current.Replace("(", "");
                    page_current = page_current.Replace(")", "");
                    string[] pags = page_current.Split(' ');

                    const int size_page = 300;
                    int last_page = int.Parse(pags[2].ToString());

                    //si tiene mas de 1 pagina entonces cargara hasta 300 comprobantes por pagina
                    if (last_page > 1)
                    {
                        driver.ExecuteScript(string.Format("let max_results = {0}; {1} {2} {3} {4} {5} {6} {7} {8}",
                              size_page,
                              "var select = document.getElementById('frmPrincipal:tablaCompEmitidos_paginator_bottom').childNodes[0];",
                              "var option = document.createElement('option');",
                              "option.text = max_results;",
                              "option.value = max_results;",
                              "select.add(option);",
                              "select.selectedIndex = select.options.length - 1;",
                              "select.dispatchEvent(new Event('change'));",
                              "select.dispatchEvent(new Event('change'));")
                              );

                        Thread.Sleep(5000);

                        //UNA VEZ QUE SE RECARGA SE DEBE VOLVER A CALCULAR LA ULTIMA PAGINA
                        page_current = driver.FindElement(By.ClassName("ui-paginator-current")).Text;
                        page_current = page_current.Replace("(", "");
                        page_current = page_current.Replace(")", "");
                        pags = page_current.Split(' ');
                        last_page = int.Parse(pags[2].ToString());
                    }

                    // Contador de numero de comprobantes
                    int i = 0;

                    //Recorre la cantidad de paginas
                    for (int current = 1; current <= last_page; current++)
                    {
                        tableComprobantes = driver.FindElement(By.Id("frmPrincipal:tablaCompEmitidos_data"));
                        rows = tableComprobantes.FindElements(By.TagName("tr")).Count;

                        if (rows > 0)
                        {
                            for (int ir = 0; ir < rows; ir++)
                            {
                                driver.ExecuteScript(string.Format("document.getElementById('frmPrincipal:tablaCompEmitidos:{0}:lnkXml').click();", i++));
                                Thread.Sleep(500);
                            }

                            if (current < last_page)
                            {
                                driver.FindElement(By.ClassName("ui-paginator-next")).Click();
                                //driver.ExecuteScript("document.getElementsByClassName('ui-paginator-next')[0].click();");
                                Thread.Sleep(7500);
                            }
                        }
                    }

                    driver.Close();
                    MessageBox.Show("Se finalizo la descarga");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error " + ex.Message);
            }
        }
    }
}
