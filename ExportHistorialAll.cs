public MemoryStream ExportHistorialCompleto()
        {
            try
            {
                ACEContext context = new();
                DataTable dt = new("Colegio");
                dt.Columns.AddRange(new DataColumn[40] {
                        new DataColumn("CODIGO_ESCUELA"),
                        new DataColumn("NOMBRE"),
                        new DataColumn("REGIONAL"),
                        new DataColumn("ESTADO"),


                        new DataColumn("SERVICIO_ID"),
                        new DataColumn("PLAN"),
                        new DataColumn("PROVEEDOR_INTERNET"),
                         new DataColumn("TIPO_SERVICIO"),
                        new DataColumn("VELOCIDAD"),
                        new DataColumn("COSTO"),
                         new DataColumn("CAUSA_DESCONEXION"),
                        new DataColumn("FECHA_DESCONEXION"),
                        new DataColumn("ESTADO_INTERNET"),
                         new DataColumn("IPV4"),
                        new DataColumn("IPV6"),
                        new DataColumn("MASCARA"),
                         new DataColumn("DNS"),
                        new DataColumn("GATEWAY"),
                        new DataColumn("NUMERO_SERVICIO"),
                         new DataColumn("UBICACCION"),
                         new DataColumn("ORDEN_COMPRA"),
                        new DataColumn("FONDOS"),


                new DataColumn("PROYECTO_ID"),
                        new DataColumn("PROYECTO"),
                        new DataColumn("NOMBRE_PROYECTO"),
                        new DataColumn("PROVEEDOR_PROYECTO"),
                        new DataColumn("TECNOLOGIA_IMPLEMENTADA"),
                        new DataColumn("FECHA_INSTALACION"),
                        new DataColumn("ESTADO_PROYECTO"),


                          new DataColumn("ELECTRICIDAD_ID"),
                        new DataColumn("ELECTRICIDAD"),
                        new DataColumn("NOMBRE_ELECTRICIDAD"),
                        new DataColumn("DESCRIPCION"),
                        new DataColumn("FUENTE_ENERGIA"),
                        new DataColumn("CAPACIDAD"),
                        new DataColumn("ESTADO_ELECTRICIDAD"),
                        new DataColumn("LICITACION"),
                        new DataColumn("TIENE_ELECTRICIDAD"),


                new DataColumn("MATRICULA"),
                        new DataColumn("PERIODO_MATRICULA")
                    });
                var centro = context.CentroEducativos.Include(r => r.ServicioInternets).ThenInclude(t => t.Plan).Include(r => r.Matriculas).Include(r => r.ServicioInternets).ThenInclude(t => t.Plan).Include(r => r.ProyectoAsignados).ThenInclude(t => t.Proyecto).Include(r => r.ElectricidadAsignada).ThenInclude(t => t.Electricidad).Include(t => t.ElectricidadAsignada).ToList();
                foreach (var colegio in centro)
                {

                    if (colegio.ServicioInternets.Count > 0 || colegio.ProyectoAsignados.Count > 0 || colegio.ElectricidadAsignada.Count > 0 || colegio.Matriculas.Count > 0)
                    {
                        int servindex = dt.Rows.Count + 1;
                        int proyindex = dt.Rows.Count + 1;
                        int elecindex = dt.Rows.Count + 1;
                        int matrindex = dt.Rows.Count + 1;

                        foreach (var internet in colegio.ServicioInternets)
                        {
                            var row = dt.Rows;
                            if (row.Count >= servindex)
                            {
                                row[servindex - 1]["CODIGO_ESCUELA"] = colegio.CentroId;
                                row[servindex - 1]["NOMBRE"] = colegio.NombreCentroEducativo;
                                row[servindex - 1]["REGIONAL"] = colegio.Regional;
                                row[servindex - 1]["ESTADO"] = colegio.Estado;
                                row[servindex - 1]["SERVICIO_ID"] = internet.ServicioId;
                                row[servindex - 1]["PLAN"] = internet.PlanId;
                                row[servindex - 1]["PROVEEDOR_INTERNET"] = internet.Plan.Proveedor;
                                row[servindex - 1]["TIPO_SERVICIO"] = internet.Plan.TipoServicio;
                                row[servindex - 1]["VELOCIDAD"] = internet.Plan.Velocidad;
                                row[servindex - 1]["COSTO"] = internet.Plan.Costo;
                                row[servindex - 1]["CAUSA_DESCONEXION"] = internet.CausaDesconexion;
                                row[servindex - 1]["FECHA_DESCONEXION"] = internet.FechaDesconexion;
                                row[servindex - 1]["ESTADO_INTERNET"] = internet.EstadoInternet;
                                row[servindex - 1]["IPV4"] = internet.Ipv4;
                                row[servindex - 1]["IPV6"] = internet.Ipv6;
                                row[servindex - 1]["MASCARA"] = internet.Mascara;
                                row[servindex - 1]["DNS"] = internet.Dns;
                                row[servindex - 1]["GATEWAY"] = internet.Gateway;
                                row[servindex - 1]["NUMERO_SERVICIO"] = internet.NumeroServicio;
                                row[servindex - 1]["UBICACCION"] = internet.Ubicaccion;
                                row[servindex - 1]["ORDEN_COMPRA"] = internet.OrdenCompra;
                                row[servindex - 1]["FONDOS"] = internet.Fondos;

                                servindex++;
                            }
                            else
                            {
                                DataRow rown = dt.NewRow();
                                rown["CODIGO_ESCUELA"] = colegio.CentroId;
                                rown["NOMBRE"] = colegio.NombreCentroEducativo;
                                rown["REGIONAL"] = colegio.Regional;
                                rown["ESTADO"] = colegio.Estado;
                                rown["SERVICIO_ID"] = internet.ServicioId;
                                rown["PLAN"] = internet.PlanId;
                                rown["PROVEEDOR_INTERNET"] = internet.Plan.Proveedor;
                                rown["TIPO_SERVICIO"] = internet.Plan.TipoServicio;
                                rown["VELOCIDAD"] = internet.Plan.Velocidad;
                                rown["COSTO"] = internet.Plan.Costo;
                                rown["CAUSA_DESCONEXION"] = internet.CausaDesconexion;
                                rown["FECHA_DESCONEXION"] = internet.FechaDesconexion;
                                rown["ESTADO_INTERNET"] = internet.EstadoInternet;
                                rown["IPV4"] = internet.Ipv4;
                                rown["IPV6"] = internet.Ipv6;
                                rown["MASCARA"] = internet.Mascara;
                                rown["DNS"] = internet.Dns;
                                rown["GATEWAY"] = internet.Gateway;
                                rown["NUMERO_SERVICIO"] = internet.NumeroServicio;
                                rown["UBICACCION"] = internet.Ubicaccion;
                                rown["ORDEN_COMPRA"] = internet.OrdenCompra;
                                rown["FONDOS"] = internet.Fondos;

                                dt.Rows.InsertAt(rown, servindex);
                                servindex++;
                            }
                        }

                        foreach (var proyecto in colegio.ProyectoAsignados)
                        {
                            var row = dt.Rows;
                            if (row.Count >= proyindex)
                            {
                                row[proyindex - 1]["CODIGO_ESCUELA"] = colegio.CentroId;
                                row[proyindex - 1]["NOMBRE"] = colegio.NombreCentroEducativo;
                                row[proyindex - 1]["REGIONAL"] = colegio.Regional;
                                row[proyindex - 1]["ESTADO"] = colegio.Estado;
                                row[proyindex - 1]["PROYECTO_ID"] = proyecto.ProyAsigId;
                                row[proyindex - 1]["PROYECTO"] = proyecto.ProyectoId;
                                row[proyindex - 1]["NOMBRE_PROYECTO"] = proyecto.Proyecto.NombreProyecto;
                                row[proyindex - 1]["PROVEEDOR_PROYECTO"] = proyecto.Proyecto.Proveedor;
                                row[proyindex - 1]["TECNOLOGIA_IMPLEMENTADA"] = proyecto.TecnologiaImplementada;
                                row[proyindex - 1]["FECHA_INSTALACION"] = proyecto.FechaInstalacion;
                                row[proyindex - 1]["ESTADO_PROYECTO"] = proyecto.Estado;
                                proyindex++;
                            }
                            else
                            {
                                DataRow rown = dt.NewRow();
                                rown["CODIGO_ESCUELA"] = colegio.CentroId;
                                rown["NOMBRE"] = colegio.NombreCentroEducativo;
                                rown["REGIONAL"] = colegio.Regional;
                                rown["ESTADO"] = colegio.Estado;
                                rown["PROYECTO_ID"] = proyecto.ProyAsigId;
                                rown["PROYECTO"] = proyecto.ProyectoId;
                                rown["NOMBRE_PROYECTO"] = proyecto.Proyecto.NombreProyecto;
                                rown["PROVEEDOR_PROYECTO"] = proyecto.Proyecto.Proveedor;
                                rown["TECNOLOGIA_IMPLEMENTADA"] = proyecto.TecnologiaImplementada;
                                rown["FECHA_INSTALACION"] = proyecto.FechaInstalacion;
                                rown["ESTADO_PROYECTO"] = proyecto.Estado;
                                dt.Rows.InsertAt(rown, proyindex);
                                proyindex++;
                            }

                        }

                        foreach (var electricidad in colegio.ElectricidadAsignada)
                        {
                            var row = dt.Rows;
                            if (row.Count >= elecindex)
                            {
                                row[elecindex - 1]["CODIGO_ESCUELA"] = colegio.CentroId;
                                row[elecindex - 1]["NOMBRE"] = colegio.NombreCentroEducativo;
                                row[elecindex - 1]["REGIONAL"] = colegio.Regional;
                                row[elecindex - 1]["ESTADO"] = colegio.Estado;
                                row[elecindex - 1]["ELECTRICIDAD_ID"] = electricidad.ElecAsigId;
                                row[elecindex - 1]["ELECTRICIDAD"] = electricidad.ElectricidadId;
                                row[elecindex - 1]["NOMBRE_ELECTRICIDAD"] = electricidad.Electricidad.NombreElectricidad;
                                row[elecindex - 1]["DESCRIPCION"] = electricidad.Electricidad.Descripcion;
                                row[elecindex - 1]["FUENTE_ENERGIA"] = electricidad.FuenteEnergia;
                                row[elecindex - 1]["CAPACIDAD"] = electricidad.Capacidad;
                                row[elecindex - 1]["ESTADO_ELECTRICIDAD"] = electricidad.Estado;
                                row[elecindex - 1]["LICITACION"] = electricidad.Licitacion;
                                row[elecindex - 1]["TIENE_ELECTRICIDAD"] = electricidad.TieneElectricidad;
                                elecindex++;
                            }
                            else
                            {
                                DataRow rown = dt.NewRow();
                                rown["CODIGO_ESCUELA"] = colegio.CentroId;
                                rown["NOMBRE"] = colegio.NombreCentroEducativo;
                                rown["REGIONAL"] = colegio.Regional;
                                rown["ESTADO"] = colegio.Estado;
                                rown["ELECTRICIDAD_ID"] = electricidad.ElecAsigId;
                                rown["ELECTRICIDAD"] = electricidad.ElectricidadId;
                                rown["NOMBRE_ELECTRICIDAD"] = electricidad.Electricidad.NombreElectricidad;
                                rown["DESCRIPCION"] = electricidad.Electricidad.Descripcion;
                                rown["FUENTE_ENERGIA"] = electricidad.FuenteEnergia;
                                rown["CAPACIDAD"] = electricidad.Capacidad;
                                rown["ESTADO_ELECTRICIDAD"] = electricidad.Estado;
                                rown["LICITACION"] = electricidad.Licitacion;
                                rown["TIENE_ELECTRICIDAD"] = electricidad.TieneElectricidad;
                                dt.Rows.InsertAt(rown, elecindex);
                                elecindex++;
                            }

                        }

                        foreach (var matricula in colegio.Matriculas)
                        {
                            var row = dt.Rows;
                            if (row.Count >= matrindex)
                            {
                                row[matrindex - 1]["CODIGO_ESCUELA"] = colegio.CentroId;
                                row[matrindex - 1]["NOMBRE"] = colegio.NombreCentroEducativo;
                                row[matrindex - 1]["REGIONAL"] = colegio.Regional;
                                row[matrindex - 1]["ESTADO"] = colegio.Estado;
                                row[matrindex - 1]["MATRICULA"] = matricula.Matriculados;
                                row[matrindex - 1]["PERIODO_MATRICULA"] = matricula.Periodo;
                                matrindex++;
                            }
                            else
                            {
                                DataRow rown = dt.NewRow();
                                rown["CODIGO_ESCUELA"] = colegio.CentroId;
                                rown["NOMBRE"] = colegio.NombreCentroEducativo;
                                rown["REGIONAL"] = colegio.Regional;
                                rown["ESTADO"] = colegio.Estado;
                                rown["MATRICULA"] = matricula.Matriculados;
                                rown["PERIODO_MATRICULA"] = matricula.Periodo; dt.Rows.InsertAt(rown, matrindex);
                                matrindex++;
                            }
                        }
                    }
                    else
                    {
                        DataRow row = dt.NewRow();
                        row["CODIGO_ESCUELA"] = colegio.CentroId;
                        row["NOMBRE"] = colegio.NombreCentroEducativo;
                        row["REGIONAL"] = colegio.Regional;
                        row["ESTADO"] = colegio.Estado;
                        dt.Rows.Add(row);
                    }
                }

                var stream = new MemoryStream();
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using var package = new ExcelPackage(stream);
                var workSheet = package.Workbook.Worksheets.Add("Sheet1");
                workSheet.Cells.LoadFromDataTable(dt, true);
                package.Save();
                return stream;
            }
            catch(Exception ex)
            {
                ErrorLogTxt(ex.ToString());
                return new MemoryStream();
            }
        }