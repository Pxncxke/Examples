public MemoryStream ExportHistorial(int select, int option, string centroid)
        {
            try
            {
                using ACEContext context = new();
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


                if (select == 1)
                {
                    if (option == 1)
                    {

                        var centro = context.CentroEducativos.Include(r => r.ServicioInternets).ThenInclude(t => t.Plan).Include(r => r.Matriculas).Include(r => r.ServicioInternets).ThenInclude(t => t.Plan).Include(r => r.ProyectoAsignados).ThenInclude(t => t.Proyecto).Include(r => r.ElectricidadAsignada).ThenInclude(t => t.Electricidad).Include(t => t.ElectricidadAsignada).Where(x => x.CentroId == centroid).ToList();
                        foreach (var colegio in centro)
                        {
                            if (colegio.ServicioInternets.Count > 0 || colegio.ProyectoAsignados.Count > 0 || colegio.ElectricidadAsignada.Count > 0 || colegio.Matriculas.Count > 0)
                            {
                                int index = 1;
                                foreach (var internet in colegio.ServicioInternets)
                                {
                                    var row = dt.Rows;
                                    if (row.Count >= index)
                                    {
                                        row[index - 1]["CODIGO_ESCUELA"] = colegio.CentroId;
                                        row[index - 1]["NOMBRE"] = colegio.NombreCentroEducativo;
                                        row[index - 1]["REGIONAL"] = colegio.Regional;
                                        row[index - 1]["ESTADO"] = colegio.Estado;
                                        row[index - 1]["SERVICIO_ID"] = internet.ServicioId;
                                        row[index - 1]["PLAN"] = internet.PlanId;
                                        row[index - 1]["PROVEEDOR_INTERNET"] = internet.Plan.Proveedor;
                                        row[index - 1]["TIPO_SERVICIO"] = internet.Plan.TipoServicio;
                                        row[index - 1]["VELOCIDAD"] = internet.Plan.Velocidad;
                                        row[index - 1]["COSTO"] = internet.Plan.Costo;
                                        row[index - 1]["CAUSA_DESCONEXION"] = internet.CausaDesconexion;
                                        row[index - 1]["FECHA_DESCONEXION"] = internet.FechaDesconexion;
                                        row[index - 1]["ESTADO_INTERNET"] = internet.EstadoInternet;
                                        row[index - 1]["IPV4"] = internet.Ipv4;
                                        row[index - 1]["IPV6"] = internet.Ipv6;
                                        row[index - 1]["MASCARA"] = internet.Mascara;
                                        row[index - 1]["DNS"] = internet.Dns;
                                        row[index - 1]["GATEWAY"] = internet.Gateway;
                                        row[index - 1]["NUMERO_SERVICIO"] = internet.NumeroServicio;
                                        row[index - 1]["UBICACCION"] = internet.Ubicaccion;
                                        row[index - 1]["ORDEN_COMPRA"] = internet.OrdenCompra;
                                        row[index - 1]["FONDOS"] = internet.Fondos;

                                        index++;
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

                                        dt.Rows.InsertAt(rown, index);
                                        index++;
                                    }
                                }
                                index = 1;
                                foreach (var proyecto in colegio.ProyectoAsignados)
                                {
                                    var row = dt.Rows;
                                    if (row.Count >= index)
                                    {
                                        row[index - 1]["CODIGO_ESCUELA"] = colegio.CentroId;
                                        row[index - 1]["NOMBRE"] = colegio.NombreCentroEducativo;
                                        row[index - 1]["REGIONAL"] = colegio.Regional;
                                        row[index - 1]["ESTADO"] = colegio.Estado;
                                        row[index - 1]["PROYECTO_ID"] = proyecto.ProyAsigId;
                                        row[index - 1]["PROYECTO"] = proyecto.ProyectoId;
                                        row[index - 1]["NOMBRE_PROYECTO"] = proyecto.Proyecto.NombreProyecto;
                                        row[index - 1]["PROVEEDOR_PROYECTO"] = proyecto.Proyecto.Proveedor;
                                        row[index - 1]["TECNOLOGIA_IMPLEMENTADA"] = proyecto.TecnologiaImplementada;
                                        row[index - 1]["FECHA_INSTALACION"] = proyecto.FechaInstalacion;
                                        row[index - 1]["ESTADO_PROYECTO"] = proyecto.Estado;
                                        index++;
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
                                        dt.Rows.InsertAt(rown, index);
                                        index++;
                                    }

                                }
                                index = 1;
                                foreach (var electricidad in colegio.ElectricidadAsignada)
                                {
                                    var row = dt.Rows;
                                    if (row.Count >= index)
                                    {
                                        row[index - 1]["CODIGO_ESCUELA"] = colegio.CentroId;
                                        row[index - 1]["NOMBRE"] = colegio.NombreCentroEducativo;
                                        row[index - 1]["REGIONAL"] = colegio.Regional;
                                        row[index - 1]["ESTADO"] = colegio.Estado;
                                        row[index - 1]["ELECTRICIDAD_ID"] = electricidad.ElecAsigId;
                                        row[index - 1]["ELECTRICIDAD"] = electricidad.ElectricidadId;
                                        row[index - 1]["NOMBRE_ELECTRICIDAD"] = electricidad.Electricidad.NombreElectricidad;
                                        row[index - 1]["DESCRIPCION"] = electricidad.Electricidad.Descripcion;
                                        row[index - 1]["FUENTE_ENERGIA"] = electricidad.FuenteEnergia;
                                        row[index - 1]["CAPACIDAD"] = electricidad.Capacidad;
                                        row[index - 1]["ESTADO_ELECTRICIDAD"] = electricidad.Estado;
                                        row[index - 1]["LICITACION"] = electricidad.Licitacion;
                                        row[index - 1]["TIENE_ELECTRICIDAD"] = electricidad.TieneElectricidad;
                                        index++;
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
                                        dt.Rows.InsertAt(rown, index);
                                        index++;
                                    }
                                }
                                index = 1;
                                foreach (var matricula in colegio.Matriculas)
                                {
                                    var row = dt.Rows;
                                    if (row.Count >= index)
                                    {
                                        row[index - 1]["CODIGO_ESCUELA"] = colegio.CentroId;
                                        row[index - 1]["NOMBRE"] = colegio.NombreCentroEducativo;
                                        row[index - 1]["REGIONAL"] = colegio.Regional;
                                        row[index - 1]["ESTADO"] = colegio.Estado;
                                        row[index - 1]["MATRICULA"] = matricula.Matriculados;
                                        row[index - 1]["PERIODO_MATRICULA"] = matricula.Periodo;
                                        index++;
                                    }
                                    else
                                    {
                                        DataRow rown = dt.NewRow();
                                        rown["CODIGO_ESCUELA"] = colegio.CentroId;
                                        rown["NOMBRE"] = colegio.NombreCentroEducativo;
                                        rown["REGIONAL"] = colegio.Regional;
                                        rown["ESTADO"] = colegio.Estado;
                                        rown["MATRICULA"] = matricula.Matriculados;
                                        rown["PERIODO_MATRICULA"] = matricula.Periodo; dt.Rows.InsertAt(rown, index);
                                        index++;
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
                    }
                    else if (option == 2)
                    {
                        var centro = context.CentroEducativos.Include(r => r.ServicioInternets).ThenInclude(t => t.Plan).Where(x => x.CentroId == centroid).ToList();

                        foreach (var colegio in centro)
                        {
                            if (colegio.ServicioInternets.Count > 0)
                            {
                                foreach (var internet in colegio.ServicioInternets)
                                {
                                    DataRow row = dt.NewRow();
                                    row["CODIGO_ESCUELA"] = colegio.CentroId;
                                    row["NOMBRE"] = colegio.NombreCentroEducativo;
                                    row["REGIONAL"] = colegio.Regional;
                                    row["ESTADO"] = colegio.Estado;
                                    row["SERVICIO_ID"] = internet.ServicioId;
                                    row["PLAN"] = internet.PlanId;
                                    row["PROVEEDOR_INTERNET"] = internet.Plan.Proveedor;
                                    row["TIPO_SERVICIO"] = internet.Plan.TipoServicio;
                                    row["VELOCIDAD"] = internet.Plan.Velocidad;
                                    row["COSTO"] = internet.Plan.Costo;
                                    row["CAUSA_DESCONEXION"] = internet.CausaDesconexion;
                                    row["FECHA_DESCONEXION"] = internet.FechaDesconexion;
                                    row["ESTADO_INTERNET"] = internet.EstadoInternet;
                                    row["IPV4"] = internet.Ipv4;
                                    row["IPV6"] = internet.Ipv6;
                                    row["MASCARA"] = internet.Mascara;
                                    row["DNS"] = internet.Dns;
                                    row["GATEWAY"] = internet.Gateway;
                                    row["NUMERO_SERVICIO"] = internet.NumeroServicio;
                                    row["UBICACCION"] = internet.Ubicaccion;
                                    row["ORDEN_COMPRA"] = internet.OrdenCompra;
                                    row["FONDOS"] = internet.Fondos;

                                    dt.Rows.Add(row);
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
                    }
                    else if (option == 3)
                    {
                        var centro = context.CentroEducativos.Include(r => r.ProyectoAsignados).ThenInclude(t => t.Proyecto).Where(x => x.CentroId == centroid).ToList();

                        foreach (var colegio in centro)
                        {
                            if (colegio.ProyectoAsignados.Count > 0)
                            {
                                foreach (var proyecto in colegio.ProyectoAsignados)
                                {
                                    DataRow row = dt.NewRow();
                                    row["CODIGO_ESCUELA"] = colegio.CentroId;
                                    row["NOMBRE"] = colegio.NombreCentroEducativo;
                                    row["REGIONAL"] = colegio.Regional;
                                    row["ESTADO"] = colegio.Estado;
                                    row["PROYECTO_ID"] = proyecto.ProyAsigId;
                                    row["PROYECTO"] = proyecto.ProyectoId;
                                    row["NOMBRE_PROYECTO"] = proyecto.Proyecto.NombreProyecto;
                                    row["PROVEEDOR_PROYECTO"] = proyecto.Proyecto.Proveedor;
                                    row["TECNOLOGIA_IMPLEMENTADA"] = proyecto.TecnologiaImplementada;
                                    row["FECHA_INSTALACION"] = proyecto.FechaInstalacion;
                                    row["ESTADO_PROYECTO"] = proyecto.Estado;
                                    dt.Rows.Add(row);
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
                    }
                    else if (option == 4)
                    {
                        var centro = context.CentroEducativos.Include(r => r.ElectricidadAsignada).ThenInclude(t => t.Electricidad).Include(t => t.ElectricidadAsignada).Where(x => x.CentroId == centroid).ToList();

                        foreach (var colegio in centro)
                        {
                            if (colegio.ElectricidadAsignada.Count > 0)
                            {
                                foreach (var electricidad in colegio.ElectricidadAsignada)
                                {
                                    DataRow row = dt.NewRow();
                                    row["CODIGO_ESCUELA"] = colegio.CentroId;
                                    row["NOMBRE"] = colegio.NombreCentroEducativo;
                                    row["REGIONAL"] = colegio.Regional;
                                    row["ESTADO"] = colegio.Estado;
                                    row["ELECTRICIDAD_ID"] = electricidad.ElecAsigId;
                                    row["ELECTRICIDAD"] = electricidad.ElectricidadId;
                                    row["NOMBRE_ELECTRICIDAD"] = electricidad.Electricidad.NombreElectricidad;
                                    row["DESCRIPCION"] = electricidad.Electricidad.Descripcion;
                                    row["FUENTE_ENERGIA"] = electricidad.FuenteEnergia;
                                    row["CAPACIDAD"] = electricidad.Capacidad;
                                    row["ESTADO_ELECTRICIDAD"] = electricidad.Estado;
                                    row["LICITACION"] = electricidad.Licitacion;
                                    row["TIENE_ELECTRICIDAD"] = electricidad.TieneElectricidad;
                                    dt.Rows.Add(row);
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
                    }
                    else if (option == 5)
                    {
                        var centro = context.CentroEducativos.Include(r => r.Matriculas).Where(x => x.CentroId == centroid).ToList();

                        foreach (var colegio in centro)
                        {
                            if (colegio.Matriculas.Count > 0)
                            {
                                foreach (var matricula in colegio.Matriculas)
                                {
                                    DataRow row = dt.NewRow();
                                    row["CODIGO_ESCUELA"] = colegio.CentroId;
                                    row["NOMBRE"] = colegio.NombreCentroEducativo;
                                    row["REGIONAL"] = colegio.Regional;
                                    row["ESTADO"] = colegio.Estado;
                                    row["MATRICULA"] = matricula.Matriculados;
                                    row["PERIODO_MATRICULA"] = matricula.Periodo;
                                    dt.Rows.Add(row);
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
                    }

                }
                else if (select == 2)
                {
                    if (option == 1)
                    {
                        var centro = context.CentroEducativos.Include(r => r.ServicioInternets).ThenInclude(t => t.Plan).Include(r => r.Matriculas).Include(r => r.ServicioInternets).ThenInclude(t => t.Plan).Include(r => r.ProyectoAsignados).ThenInclude(t => t.Proyecto).Include(r => r.ElectricidadAsignada).ThenInclude(t => t.Electricidad).Include(t => t.ElectricidadAsignada).Where(x => x.NombreCentroEducativo.Contains(centroid)).ToList();
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
                    }
                    else if (option == 2)
                    {
                        var centro = context.CentroEducativos.Include(r => r.ServicioInternets).ThenInclude(t => t.Plan).Where(x => x.NombreCentroEducativo.Contains(centroid)).ToList();

                        foreach (var colegio in centro)
                        {
                            if (colegio.ServicioInternets.Count > 0)
                            {
                                foreach (var internet in colegio.ServicioInternets)
                                {
                                    DataRow row = dt.NewRow();
                                    row["CODIGO_ESCUELA"] = colegio.CentroId;
                                    row["NOMBRE"] = colegio.NombreCentroEducativo;
                                    row["REGIONAL"] = colegio.Regional;
                                    row["ESTADO"] = colegio.Estado;
                                    row["SERVICIO_ID"] = internet.ServicioId;
                                    row["PLAN"] = internet.PlanId;
                                    row["PROVEEDOR_INTERNET"] = internet.Plan.Proveedor;
                                    row["TIPO_SERVICIO"] = internet.Plan.TipoServicio;
                                    row["VELOCIDAD"] = internet.Plan.Velocidad;
                                    row["COSTO"] = internet.Plan.Costo;
                                    row["CAUSA_DESCONEXION"] = internet.CausaDesconexion;
                                    row["FECHA_DESCONEXION"] = internet.FechaDesconexion;
                                    row["ESTADO_INTERNET"] = internet.EstadoInternet;
                                    row["IPV4"] = internet.Ipv4;
                                    row["IPV6"] = internet.Ipv6;
                                    row["MASCARA"] = internet.Mascara;
                                    row["DNS"] = internet.Dns;
                                    row["GATEWAY"] = internet.Gateway;
                                    row["NUMERO_SERVICIO"] = internet.NumeroServicio;
                                    row["UBICACCION"] = internet.Ubicaccion;
                                    row["ORDEN_COMPRA"] = internet.OrdenCompra;
                                    row["FONDOS"] = internet.Fondos;

                                    dt.Rows.Add(row);
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
                    }
                    else if (option == 3)
                    {
                        var centro = context.CentroEducativos.Include(r => r.ProyectoAsignados).ThenInclude(t => t.Proyecto).Where(x => x.NombreCentroEducativo.Contains(centroid)).ToList();

                        foreach (var colegio in centro)
                        {
                            if (colegio.ProyectoAsignados.Count > 0)
                            {
                                foreach (var proyecto in colegio.ProyectoAsignados)
                                {
                                    DataRow row = dt.NewRow();
                                    row["CODIGO_ESCUELA"] = colegio.CentroId;
                                    row["NOMBRE"] = colegio.NombreCentroEducativo;
                                    row["REGIONAL"] = colegio.Regional;
                                    row["ESTADO"] = colegio.Estado;
                                    row["PROYECTO_ID"] = proyecto.ProyAsigId;
                                    row["PROYECTO"] = proyecto.ProyectoId;
                                    row["NOMBRE_PROYECTO"] = proyecto.Proyecto.NombreProyecto;
                                    row["PROVEEDOR_PROYECTO"] = proyecto.Proyecto.Proveedor;
                                    row["TECNOLOGIA_IMPLEMENTADA"] = proyecto.TecnologiaImplementada;
                                    row["FECHA_INSTALACION"] = proyecto.FechaInstalacion;
                                    row["ESTADO_PROYECTO"] = proyecto.Estado;
                                    dt.Rows.Add(row);
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
                    }
                    else if (option == 4)
                    {
                        var centro = context.CentroEducativos.Include(r => r.ElectricidadAsignada).ThenInclude(t => t.Electricidad).Include(t => t.ElectricidadAsignada).Where(x => x.NombreCentroEducativo.Contains(centroid)).ToList();

                        foreach (var colegio in centro)
                        {
                            if (colegio.ElectricidadAsignada.Count > 0)
                            {
                                foreach (var electricidad in colegio.ElectricidadAsignada)
                                {
                                    DataRow row = dt.NewRow();
                                    row["CODIGO_ESCUELA"] = colegio.CentroId;
                                    row["NOMBRE"] = colegio.NombreCentroEducativo;
                                    row["REGIONAL"] = colegio.Regional;
                                    row["ESTADO"] = colegio.Estado;
                                    row["ELECTRICIDAD_ID"] = electricidad.ElecAsigId;
                                    row["ELECTRICIDAD"] = electricidad.ElectricidadId;
                                    row["NOMBRE_ELECTRICIDAD"] = electricidad.Electricidad.NombreElectricidad;
                                    row["DESCRIPCION"] = electricidad.Electricidad.Descripcion;
                                    row["FUENTE_ENERGIA"] = electricidad.FuenteEnergia;
                                    row["CAPACIDAD"] = electricidad.Capacidad;
                                    row["ESTADO_ELECTRICIDAD"] = electricidad.Estado;
                                    row["LICITACION"] = electricidad.Licitacion;
                                    row["TIENE_ELECTRICIDAD"] = electricidad.TieneElectricidad;
                                    dt.Rows.Add(row);
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
                    }
                    else if (option == 5)
                    {
                        var centro = context.CentroEducativos.Include(r => r.Matriculas).Where(x => x.NombreCentroEducativo.Contains(centroid)).ToList();

                        foreach (var colegio in centro)
                        {
                            if (colegio.Matriculas.Count > 0)
                            {
                                foreach (var matricula in colegio.Matriculas)
                                {
                                    DataRow row = dt.NewRow();
                                    row["CODIGO_ESCUELA"] = colegio.CentroId;
                                    row["NOMBRE"] = colegio.NombreCentroEducativo;
                                    row["REGIONAL"] = colegio.Regional;
                                    row["ESTADO"] = colegio.Estado;
                                    row["MATRICULA"] = matricula.Matriculados;
                                    row["PERIODO_MATRICULA"] = matricula.Periodo;
                                    dt.Rows.Add(row);
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
            catch (Exception ex)
            {
                ErrorLogTxt(ex.ToString());
                return new MemoryStream();
            }
        }