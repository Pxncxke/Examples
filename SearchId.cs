public string BuscarTribunal(TopModel topModel) 
        {
            using (VerificacionIdentidadSoapClient TE = new VerificacionIdentidadSoapClient())
            {

                Llaves_Webservice cert = new Llaves_Webservice();
                using (MeducaWSEntities meducaWSEntities = new MeducaWSEntities())
                {
                    cert = meducaWSEntities.Llaves_Webservice.Where(x => x.FechaExpiracion >= DateTime.Now).FirstOrDefault();
                    if (cert == null)
                    {
                        meducaWSEntities.Registro_Consulta_INSERT(topModel.AppId, topModel.Cedula, DateTime.Now, "Error de certificado", false);
                        return Security.EnviarMensajeError("Error de certificado");
                    }
                    var certbytes = Encoding.Default.GetBytes(cert.Llave);
                    TE.ClientCredentials.ClientCertificate.Certificate = new X509Certificate2(certbytes);
                }

                AumentarContador(DateTime.Now, cert.LlaveId);

                var result = TE.VerificarPersona(topModel.Cedula);

                var nPersona = new Persona
                {
                    DIFUNTO = false
                };
                var doc = new XmlDocument();
                doc.LoadXml(result);
                var nsmgr = new XmlNamespaceManager(doc.NameTable);
                nsmgr.AddNamespace("a", "http://tempuri.org/DatasetPersona.xsd");

                var mensajes = doc.SelectNodes("//a:Mensajes", nsmgr);

                using (MeducaWSEntities meducaWSEntities = new MeducaWSEntities()) 
                {
                    foreach (XmlNode xn in mensajes)
                    {
                        if (xn["CodMensaje"] != null && xn["Mensaje"] != null)
                        {
                            if(xn["CodMensaje"].InnerText == "530") 
                            {
                                meducaWSEntities.Registro_Consulta_INSERT(topModel.AppId, topModel.Cedula, DateTime.Now, xn["Mensaje"].InnerText, false);
                                return Security.EnviarMensajeError(xn["Mensaje"].InnerText);
                            }
                            nPersona.ID_COMENTARIO = xn["CodMensaje"].InnerText;
                            nPersona.COMENTARIO = xn["Mensaje"].InnerText;
                            if (xn["CodMensaje"].InnerText == "534") 
                            {
                                nPersona.DIFUNTO = true;
                            }
                        }
                    }
                    meducaWSEntities.Registro_Consulta_INSERT(topModel.AppId, topModel.Cedula, DateTime.Now, nPersona.COMENTARIO, false);
                }


                var pPublica = doc.SelectNodes("//a:PersonaPublica", nsmgr);

                foreach (XmlNode xn in pPublica)
                {
                    if (xn["cedula"] != null)
                        nPersona.CEDULA = xn["cedula"].InnerText;
                    if (xn["primer_nombre"] != null)
                        nPersona.PRIMER_NOMBRE = xn["primer_nombre"].InnerText;
                    if (xn["segundo_nombre"] != null)
                        nPersona.SEGUNDO_NOMBRE = xn["segundo_nombre"].InnerText;
                    if (xn["apellido_paterno"] != null)
                        nPersona.APELLIDO_PATERNO = xn["apellido_paterno"].InnerText;
                    if (xn["apellido_materno"] != null)
                        nPersona.APELLIDO_MATERNO = xn["apellido_materno"].InnerText;
                    if (xn["fecha_nacimiento"] != null)
                        nPersona.FECHA_NACIMIENTO = xn["fecha_nacimiento"].InnerText;
                    if (xn["sexo"] != null)
                        nPersona.SEXO = xn["sexo"].InnerText;
                    if (xn["estado_civil"] != null)
                        nPersona.ESTADO_CIVIL = xn["estado_civil"].InnerText;
                    if (xn["pais"] != null)
                        nPersona.PAIS = xn["pais"].InnerText;
                    if (xn["prov"] != null)
                        nPersona.PROV = xn["prov"].InnerText;
                    if (xn["distrito"] != null)
                        nPersona.DISTRITO = xn["distrito"].InnerText;
                    if (xn["corregimiento"] != null)
                        nPersona.CORREGIMIENTO = xn["corregimiento"].InnerText;
                    if (xn["fecha_vencimiento_cedula"] != null)
                        nPersona.FECHA_VENCIMIENTO_CEDULA = xn["fecha_vencimiento_cedula"].InnerText;
                    if (xn["lugarnacimientope"] != null)
                        nPersona.LUGARNACIMIENTOPE = xn["lugarnacimientope"].InnerText;
                    if (xn["LugarDeNacimiento"] != null)
                        nPersona.LUGARDENACIMIENTO = xn["LugarDeNacimiento"].InnerText;
                    if (xn["barrio_residencia"] != null)
                        nPersona.BARRIO_RESIDENCIA = xn["barrio_residencia"].InnerText;
                    if (xn["edificio_casa"] != null)
                        nPersona.EDIFICIO_CASA = xn["edificio_casa"].InnerText;
                    if (xn["nombreCedula"] != null)
                        nPersona.NOMBRE_CEDULA = xn["nombreCedula"].InnerText;
                }

                var pConfidencial = doc.SelectNodes("//a:PersonaConfidencial", nsmgr);

                foreach (XmlNode xn in pConfidencial)
                {
                    if (xn["primer_nombre_madre"] != null)
                        nPersona.PRIMER_NOMBRE_MADRE = xn["primer_nombre_madre"].InnerText;
                    if(xn["segundo_nombre_madre"] != null)
                        nPersona.SEGUNDO_NOMBRE_MADRE = xn["segundo_nombre_madre"].InnerText;
                    if (xn["apellido_paterno_madre"] != null)
                        nPersona.APELLIDO_PATERNO_MADRE = xn["apellido_paterno_madre"].InnerText;
                    if (xn["apellido_materno_madre"] != null)
                        nPersona.APELLIDO_MATERNO_MADRE = xn["apellido_materno_madre"].InnerText;
                    if (xn["apellido_casada_madre"] != null)
                        nPersona.APELLIDO_CASADA_MADRE = xn["apellido_casada_madre"].InnerText;
                    if (xn["primer_nombre_padre"] != null)
                        nPersona.PRIMER_NOMBRE_PADRE = xn["primer_nombre_padre"].InnerText;
                    if (xn["segundo_nombre_padre"] != null)
                        nPersona.SEGUNDO_NOMBRE_PADRE = xn["segundo_nombre_padre"].InnerText;
                    if (xn["apellido_paterno_padre"] != null)
                        nPersona.APELLIDO_PATERNO_PADRE = xn["apellido_paterno_padre"].InnerText;
                    if (xn["apellido_materno_padre"] != null)
                        nPersona.APELLIDO_MATERNO_PADRE = xn["apellido_materno_padre"].InnerText;
                    if (xn["nombre_centro"] != null)
                        nPersona.NOMBRE_CENTRO = xn["nombre_centro"].InnerText;
                    if (xn["provincia_nombre"] != null)
                        nPersona.PROVINCIA_NOMBRE = xn["provincia_nombre"].InnerText;
                    if (xn["distrito_nombre"] != null)
                        nPersona.DISTRITO_NOMBRE = xn["distrito_nombre"].InnerText;
                    if (xn["corregimiento_nombre"] != null)
                        nPersona.CORREGIMIENTRO_NOMBRE = xn["corregimiento_nombre"].InnerText;
                    if (xn["cedula_madre"] != null)
                        nPersona.CEDULA_MADRE = xn["cedula_madre"].InnerText;
                    if (xn["cedula_padre"] != null)
                        nPersona.CEDULA_PADRE = xn["cedula_padre"].InnerText;
                }

                var imagenes = doc.SelectNodes("//a:Imagenes", nsmgr);

                foreach (XmlNode xn in imagenes)
                {
                    if (xn["UrlFoto"] != null)
                        nPersona.URLFOTO = xn["UrlFoto"].InnerText;
                    if (xn["UrlFirma"] != null)
                        nPersona.URLFIRMA = xn["UrlFirma"].InnerText;
                }

                nPersona.ULTIMA_MODIFICACION = DateTime.Now;

                topModel.Persona = nPersona;

                return InsertarPersona(topModel);
            }
        }