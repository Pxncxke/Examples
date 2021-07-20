 public async Task<bool> Import(HttpPostedFileBase postedFile, string path, string user, int cargapago)
        {
            try
            {
                using (SCDPEntities Db = new SCDPEntities())
                {
                    var colegio = Db.Colegios_Asignados.Include(r => r.Colegio).Where(x => x.USUARIO_ID == user).FirstOrDefault();
                    int count = 0;
                    if (colegio.Colegio.PASE_U)
                    {
                        string filePath = string.Empty;
                        var name = await Cipher.GenerateRandomCode(10);
                        string extension = Path.GetExtension(postedFile.FileName);
                        var ext = DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + colegio.COLEGIO_ID + user + name + extension;
                        string conString = conString = ConfigurationManager.ConnectionStrings["Excel07ConString"].ConnectionString;
                        int stx;
                        filePath = path + Path.GetFileName(ext);
                        postedFile.SaveAs(filePath);

                        switch (extension)
                        {
                            case ".xls": For Excel 97-03.  
                                conString = ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;
                                break;
                            case ".xlsx": For Excel 07 and above.  
                                conString = ConfigurationManager.ConnectionStrings["Excel07ConString"].ConnectionString;
                                break;
                        }


                        DataTable dt = new DataTable();
                        conString = string.Format(conString, filePath);

                        using (OleDbConnection connExcel = new OleDbConnection(conString))
                        {
                            using (OleDbCommand cmdExcel = new OleDbCommand())
                            {
                                using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                                {
                                    cmdExcel.Connection = connExcel;

                                   
                                    connExcel.Open();
                                    DataTable dtExcelSchema;
                                    dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                                    string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                                    connExcel.Close();


                                    connExcel.Open();
                                    cmdExcel.CommandText = "SELECT * From [" + sheetName + "]";
                                    odaExcel.SelectCommand = cmdExcel;
                                    odaExcel.Fill(dt);
                                    connExcel.Close();
                                }
                            }
                        }

                        DataColumn column = new DataColumn("FECHA", typeof(string))
                        {
                            DefaultValue = DateTime.Now
                        };

                        dt.Columns.Add(column);

                        column = new DataColumn("USUARIO", typeof(string))
                        {
                            DefaultValue = user
                        };

                        dt.Columns.Add(column);

                        if (cargapago == 1)
                        {
                            conString = ConfigurationManager.ConnectionStrings["DBCS"].ConnectionString;
                            using (SqlConnection con = new SqlConnection(conString))
                            {
                                using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                                {
                                    
                                    sqlBulkCopy.DestinationTableName = "CargaEstudiantes";
                                    sqlBulkCopy.ColumnMappings.Add(1, "CEDULA");
                                    sqlBulkCopy.ColumnMappings.Add(2, "PRIMER_NOMBRE");
                                    sqlBulkCopy.ColumnMappings.Add(3, "SEGUNDO_NOMBRE");
                                    sqlBulkCopy.ColumnMappings.Add(4, "APELLIDO_PATERNO");
                                    sqlBulkCopy.ColumnMappings.Add(5, "APELLIDO_MATERNO");
                                    sqlBulkCopy.ColumnMappings.Add(6, "PASAPORTE");
                                    sqlBulkCopy.ColumnMappings.Add(7, "GRADO");
                                    sqlBulkCopy.ColumnMappings.Add(8, "GRUPO");
                                    sqlBulkCopy.ColumnMappings.Add(9, "BACHILLERATO");
                                    sqlBulkCopy.ColumnMappings.Add(10, "FECHA_NACIMIENTO");
                                    sqlBulkCopy.ColumnMappings.Add(11, "CEDULA_ACUDIENTE");
                                    sqlBulkCopy.ColumnMappings.Add(12, "NOMBRE_COMPLETO_ACUDIENTE");
                                    sqlBulkCopy.ColumnMappings.Add(13, "PASAPORTE_ACUDIENTE");
                                    sqlBulkCopy.ColumnMappings.Add(14, "FECHA");
                                    sqlBulkCopy.ColumnMappings.Add(15, "USUARIO");
                                    con.Open();
                                    sqlBulkCopy.WriteToServer(dt);
                                    con.Close();
                                }
                            }

                            var carga = Db.CargaEstudiantes.Where(x => x.USUARIO == user).ToList();
                          

                            var rx = new Regex(@"^(PE|E|N|[23456789](?:AV|PI)?|1[0123]?(?:AV|PI)?)-(?:[1-9]|[1-9][0-9]{1,3})-(?:[1-9]|[1-9][0-9]{1,5})$", RegexOptions.IgnoreCase);

                            foreach (var nuevo in carga)
                            {
                                if (!String.IsNullOrWhiteSpace(nuevo.CEDULA_ACUDIENTE))
                                {
                                    nuevo.CEDULA_ACUDIENTE = nuevo.CEDULA_ACUDIENTE.ToUpper().Trim();

                                    if (!rx.IsMatch(nuevo.CEDULA_ACUDIENTE))
                                    {
                                        continue;
                                    }
                                }
                                else if (String.IsNullOrWhiteSpace(nuevo.PASAPORTE_ACUDIENTE))
                                {
                                    continue;
                                }
                                if (String.IsNullOrWhiteSpace(nuevo.FECHA_NACIMIENTO) || String.IsNullOrWhiteSpace(nuevo.GRADO) || String.IsNullOrWhiteSpace(nuevo.NOMBRE_COMPLETO_ACUDIENTE))
                                {
                                    continue;
                                }
                                nuevo.GRADO = nuevo.GRADO.ToUpper().Trim();

                                if (!String.IsNullOrWhiteSpace(nuevo.CEDULA))
                                {
                                    nuevo.CEDULA = nuevo.CEDULA.ToUpper().Trim();

                                    if (rx.IsMatch(nuevo.CEDULA))
                                    {
                                        var valido = Db.Estudiantes.Where(x => x.CEDULA == nuevo.CEDULA).FirstOrDefault();
                                        var tribunal = await ValidarEstudiantes(nuevo.CEDULA);
                                        if (tribunal.MensajeError == null)
                                        {
                                            if (valido == null)
                                            {
                                                var student = new Estudiante();

                                                var resultDate = new DateTime();
                                                if (DateTime.TryParse(tribunal.Persona.FECHA_NACIMIENTO, out resultDate))
                                                {
                                                    int edad = DateTime.Today.Year - resultDate.Date.Year;

                                                    if (resultDate.Date > DateTime.Today.AddYears(-edad)) edad--;

                                                    if (edad <= 21 && edad >= 4)
                                                    {
                                                        student.FECHA_NACIMIENTO = resultDate.Month + "-" + resultDate.Day + "-" + resultDate.Year;
                                                    }
                                                    else
                                                    {
                                                        continue;
                                                    }
                                                }
                                                else
                                                {
                                                    continue;
                                                }
                                                if (nuevo.GRADO.Contains("°"))
                                                {
                                                    nuevo.GRADO = nuevo.GRADO.Replace("°", string.Empty);
                                                }
                                                if (nuevo.GRADO.Length <= 2 && int.TryParse(nuevo.GRADO, out stx))
                                                {
                                                    student.GRADO = nuevo.GRADO;
                                                }
                                                else
                                                {
                                                    continue;
                                                }

                                                student.CEDULA = nuevo.CEDULA;
                                                student.COLEGIO_ID = colegio.COLEGIO_ID;
                                                student.PRIMER_NOMBRE = tribunal.Persona.PRIMER_NOMBRE;
                                                student.SEGUNDO_NOMBRE = tribunal.Persona.SEGUNDO_NOMBRE;
                                                student.APELLIDO_PATERNO = tribunal.Persona.APELLIDO_PATERNO;
                                                student.APELLIDO_MATERNO = tribunal.Persona.APELLIDO_MATERNO;
                                                student.PASAPORTE = string.Empty;
                                                student.CEDULA_ACUDIENTE = nuevo.CEDULA_ACUDIENTE;
                                                student.NOMBRE_ACUDIENTE = nuevo.NOMBRE_COMPLETO_ACUDIENTE;
                                                student.PASAPORTE_ACUDIENTE = nuevo.PASAPORTE_ACUDIENTE;
                                                student.FECHA_ULTIMA_MODIFICACION = DateTime.Now;
                                                student.PERIODO = DateTime.Now.Year;



                                                if (!String.IsNullOrWhiteSpace(nuevo.GRUPO))
                                                {
                                                    if (nuevo.GRUPO.Length <= 4)
                                                    {
                                                        student.GRUPO = nuevo.GRUPO;
                                                    }
                                                }

                                                if (nuevo.BACHILLERATO != null)
                                                
                                                {
                                                    if (nuevo.BACHILLERATO.Length <= 50)
                                                    {
                                                        student.BACHILLERATO = nuevo.BACHILLERATO;
                                                    }
                                                }

                                                Db.Estudiantes.Add(student);
                                                Db.CargaEstudiantes.Remove(nuevo);
                                                await Db.SaveChangesAsync();
                                                count++;
                                            }
                                            else
                                            {

                                                var resultDate = new DateTime();
                                                if (DateTime.TryParse(tribunal.Persona.FECHA_NACIMIENTO, out resultDate))
                                                {
                                                    int edad = DateTime.Today.Year - resultDate.Date.Year;

                                                    if (resultDate.Date > DateTime.Today.AddYears(-edad)) edad--;

                                                    if (edad <= 21 && edad >= 4)
                                                    {
                                                        valido.FECHA_NACIMIENTO = resultDate.Month + "-" + resultDate.Day + "-" + resultDate.Year;
                                                    }
                                                    else
                                                    {
                                                        continue;
                                                    }
                                                }
                                                else
                                                {
                                                    continue;
                                                }
                                                if (nuevo.GRADO.Contains("°"))
                                                {
                                                    nuevo.GRADO = nuevo.GRADO.Replace("°", string.Empty);
                                                }
                                                if (nuevo.GRADO.Length <= 2 && int.TryParse(nuevo.GRADO, out stx))
                                                {
                                                    valido.GRADO = nuevo.GRADO;
                                                }
                                                else
                                                {
                                                    continue;
                                                }

                                                valido.COLEGIO_ID = colegio.COLEGIO_ID;

                                                valido.PRIMER_NOMBRE = tribunal.Persona.PRIMER_NOMBRE;
                                                valido.SEGUNDO_NOMBRE = tribunal.Persona.SEGUNDO_NOMBRE;
                                                valido.APELLIDO_PATERNO = tribunal.Persona.APELLIDO_PATERNO;
                                                valido.APELLIDO_MATERNO = tribunal.Persona.APELLIDO_MATERNO;

                                                valido.CEDULA_ACUDIENTE = nuevo.CEDULA_ACUDIENTE;
                                                valido.NOMBRE_ACUDIENTE = nuevo.NOMBRE_COMPLETO_ACUDIENTE;
                                                valido.PASAPORTE_ACUDIENTE = nuevo.PASAPORTE_ACUDIENTE;
                                                valido.FECHA_ULTIMA_MODIFICACION = DateTime.Now;
                                                valido.PERIODO = DateTime.Now.Year;

                                                if (!String.IsNullOrWhiteSpace(nuevo.GRUPO))
                                                {
                                                    if (nuevo.GRUPO.Length <= 4)
                                                    {
                                                        valido.GRUPO = nuevo.GRUPO;
                                                    }
                                                }
                                                if (nuevo.BACHILLERATO != null)
                                               
                                                {
                                                    if (nuevo.BACHILLERATO.Length <= 50)
                                                    {
                                                        valido.BACHILLERATO = nuevo.BACHILLERATO;
                                                    }
                                                }


                                                Db.CargaEstudiantes.Remove(nuevo);
                                                await Db.SaveChangesAsync();
                                                count++;
                                            }
                                        }

                                    }
                                    else
                                    {
                                        Db.CargaEstudiantes.Remove(nuevo);
                                        await Db.SaveChangesAsync();
                                    }
                                }
                                else if (!String.IsNullOrWhiteSpace(nuevo.PASAPORTE))
                                {
                                    var estudiante = Db.Estudiantes.Where(x => x.PASAPORTE == nuevo.PASAPORTE).FirstOrDefault();

                                    if (estudiante == null)
                                    {
                                        var student = new Estudiante();
                                        var resultDate = new DateTime();

                                        if (DateTime.TryParse(nuevo.FECHA_NACIMIENTO, out resultDate))
                                        {
                                            int edad = DateTime.Today.Year - resultDate.Date.Year;

                                            if (resultDate.Date > DateTime.Today.AddYears(-edad)) edad--;

                                            if (edad <= 21 && edad >= 4)
                                            {
                                                student.FECHA_NACIMIENTO = resultDate.Month + "-" + resultDate.Day + "-" + resultDate.Year;
                                            }
                                            else
                                            {
                                                continue;
                                            }
                                        }
                                        else
                                        {
                                            continue;
                                        }

                                        if (nuevo.GRADO.Contains("°"))
                                        {
                                            nuevo.GRADO = nuevo.GRADO.Replace("°", string.Empty);
                                        }
                                        if (nuevo.GRADO.Length <= 2 && int.TryParse(nuevo.GRADO, out stx))
                                        {
                                            student.GRADO = nuevo.GRADO;
                                        }
                                        else
                                        {
                                            continue;
                                        }




                                        student.CEDULA = null;
                                        student.COLEGIO_ID = colegio.COLEGIO_ID;
                                        student.PRIMER_NOMBRE = nuevo.PRIMER_NOMBRE;
                                        student.SEGUNDO_NOMBRE = nuevo.SEGUNDO_NOMBRE;
                                        student.APELLIDO_PATERNO = nuevo.APELLIDO_PATERNO;
                                        student.APELLIDO_MATERNO = nuevo.APELLIDO_MATERNO;

                                        student.PASAPORTE = nuevo.PASAPORTE;
                                        student.CEDULA_ACUDIENTE = nuevo.CEDULA_ACUDIENTE;
                                        student.NOMBRE_ACUDIENTE = nuevo.NOMBRE_COMPLETO_ACUDIENTE;
                                        student.PASAPORTE_ACUDIENTE = nuevo.PASAPORTE_ACUDIENTE;
                                        student.FECHA_ULTIMA_MODIFICACION = DateTime.Now;
                                        student.PERIODO = DateTime.Now.Year;




                                        if (!String.IsNullOrWhiteSpace(nuevo.GRUPO))
                                        {
                                            if (nuevo.GRUPO.Length <= 4)
                                            {
                                                student.GRUPO = nuevo.GRUPO;
                                            }
                                        }
                                        if (nuevo.BACHILLERATO != null)
                                       
                                        {
                                            if (nuevo.BACHILLERATO.Length <= 50)
                                            {
                                                student.BACHILLERATO = nuevo.BACHILLERATO;
                                            }
                                        }

                                        Db.Estudiantes.Add(student);
                                        Db.CargaEstudiantes.Remove(nuevo);
                                        await Db.SaveChangesAsync();
                                        count++;
                                    }
                                    else
                                    {
                                        var resultDate = new DateTime();
                                        if (DateTime.TryParse(nuevo.FECHA_NACIMIENTO, out resultDate))
                                        {
                                            int edad = DateTime.Today.Year - resultDate.Date.Year;

                                            if (resultDate.Date > DateTime.Today.AddYears(-edad)) edad--;

                                            if (edad <= 21 && edad >= 4)
                                            {
                                                estudiante.FECHA_NACIMIENTO = resultDate.Month + "-" + resultDate.Day + "-" + resultDate.Year;
                                            }
                                            else
                                            {
                                                continue;
                                            }
                                        }
                                        else
                                        {
                                            continue;
                                        }
                                        if (nuevo.GRADO.Contains("°"))
                                        {
                                            nuevo.GRADO = nuevo.GRADO.Replace("°", string.Empty);
                                        }
                                        if (nuevo.GRADO.Length <= 2 && int.TryParse(nuevo.GRADO, out stx))
                                        {
                                            estudiante.GRADO = nuevo.GRADO;
                                        }
                                        else
                                        {
                                            continue;
                                        }

                                        estudiante.COLEGIO_ID = colegio.COLEGIO_ID;

                                        estudiante.PRIMER_NOMBRE = nuevo.PRIMER_NOMBRE;
                                        estudiante.SEGUNDO_NOMBRE = nuevo.SEGUNDO_NOMBRE;
                                        estudiante.APELLIDO_PATERNO = nuevo.APELLIDO_PATERNO;
                                        estudiante.APELLIDO_MATERNO = nuevo.APELLIDO_MATERNO;

                                        estudiante.CEDULA_ACUDIENTE = nuevo.CEDULA_ACUDIENTE;
                                        estudiante.NOMBRE_ACUDIENTE = nuevo.NOMBRE_COMPLETO_ACUDIENTE;
                                        estudiante.PASAPORTE_ACUDIENTE = nuevo.PASAPORTE_ACUDIENTE;
                                        estudiante.FECHA_ULTIMA_MODIFICACION = DateTime.Now;
                                        estudiante.PERIODO = DateTime.Now.Year;



                                        if (!String.IsNullOrWhiteSpace(nuevo.GRUPO))
                                        {
                                            if (nuevo.GRUPO.Length <= 4)
                                            {
                                                estudiante.GRUPO = nuevo.GRUPO;
                                            }
                                        }
                                        if (nuevo.BACHILLERATO != null)
                                       
                                        {
                                            if (nuevo.BACHILLERATO.Length <= 50)
                                            {
                                                estudiante.BACHILLERATO = nuevo.BACHILLERATO;
                                            }
                                        }


                                        Db.CargaEstudiantes.Remove(nuevo);
                                        await Db.SaveChangesAsync();
                                        count++;
                                    }
                                }
                            }
                        }
                        else if (cargapago == 2)
                        {
                            conString = ConfigurationManager.ConnectionStrings["DBCS"].ConnectionString;
                            using (SqlConnection con = new SqlConnection(conString))
                            {
                                using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                                {

                                    sqlBulkCopy.DestinationTableName = "CargaEstudiantesAsistencia";
                                    sqlBulkCopy.ColumnMappings.Add(1, "CEDULA");
                                    sqlBulkCopy.ColumnMappings.Add(2, "PRIMER_NOMBRE");
                                    sqlBulkCopy.ColumnMappings.Add(3, "SEGUNDO_NOMBRE");
                                    sqlBulkCopy.ColumnMappings.Add(4, "APELLIDO_PATERNO");
                                    sqlBulkCopy.ColumnMappings.Add(5, "APELLIDO_MATERNO");
                                    sqlBulkCopy.ColumnMappings.Add(6, "PASAPORTE");
                                    sqlBulkCopy.ColumnMappings.Add(7, "GRADO");
                                    sqlBulkCopy.ColumnMappings.Add(8, "GRUPO");
                                    sqlBulkCopy.ColumnMappings.Add(9, "BACHILLERATO");
                                    sqlBulkCopy.ColumnMappings.Add(10, "FECHA_NACIMIENTO");
                                    sqlBulkCopy.ColumnMappings.Add(11, "CEDULA_ACUDIENTE");
                                    sqlBulkCopy.ColumnMappings.Add(12, "NOMBRE_COMPLETO_ACUDIENTE");
                                    sqlBulkCopy.ColumnMappings.Add(13, "PASAPORTE_ACUDIENTE");
                                    sqlBulkCopy.ColumnMappings.Add(14, "ASISTIO");
                                    sqlBulkCopy.ColumnMappings.Add(15, "FECHA");
                                    sqlBulkCopy.ColumnMappings.Add(16, "USUARIO");
                                    con.Open();
                                    sqlBulkCopy.WriteToServer(dt);
                                    con.Close();
                                }
                            }

                            var carga = Db.CargaEstudiantesAsistencias.Where(x => x.USUARIO == user).ToList();
                           

                            var rx = new Regex(@"^(PE|E|N|[23456789](?:AV|PI)?|1[0123]?(?:AV|PI)?)-(?:[1-9]|[1-9][0-9]{1,3})-(?:[1-9]|[1-9][0-9]{1,5})$", RegexOptions.IgnoreCase);

                            foreach (var nuevo in carga)
                            {
                                if (!String.IsNullOrWhiteSpace(nuevo.CEDULA_ACUDIENTE))
                                {
                                    nuevo.CEDULA_ACUDIENTE =  nuevo.CEDULA_ACUDIENTE.ToUpper().Trim();

                                    if (!rx.IsMatch(nuevo.CEDULA_ACUDIENTE))
                                    {
                                        continue;
                                    }
                                }
                                else if (String.IsNullOrWhiteSpace(nuevo.PASAPORTE_ACUDIENTE))
                                {
                                    continue;
                                }
                                if (String.IsNullOrWhiteSpace(nuevo.FECHA_NACIMIENTO) || String.IsNullOrWhiteSpace(nuevo.GRADO) || String.IsNullOrWhiteSpace(nuevo.NOMBRE_COMPLETO_ACUDIENTE))
                                {
                                    continue;
                                }
                                nuevo.GRADO = nuevo.GRADO.ToUpper().Trim();

                                if (!String.IsNullOrWhiteSpace(nuevo.CEDULA))
                                {
                                    nuevo.CEDULA = nuevo.CEDULA.ToUpper().Trim();

                                    if (rx.IsMatch(nuevo.CEDULA))
                                    {
                                        var valido = Db.Estudiantes.Where(x => x.CEDULA == nuevo.CEDULA).FirstOrDefault();
                                        var tribunal = await ValidarEstudiantes(nuevo.CEDULA);
                                        if (tribunal.MensajeError == null)
                                        {
                                            if (valido == null)
                                            {
                                                var student = new Estudiante();

                                                var resultDate = new DateTime();
                                                if (DateTime.TryParse(tribunal.Persona.FECHA_NACIMIENTO, out resultDate))
                                                {
                                                    int edad = DateTime.Today.Year - resultDate.Date.Year;

                                                    if (resultDate.Date > DateTime.Today.AddYears(-edad)) edad--;

                                                    if (edad <= 21 && edad >= 4)
                                                    {
                                                        student.FECHA_NACIMIENTO = resultDate.Month + "-" + resultDate.Day + "-" + resultDate.Year;
                                                    }
                                                    else
                                                    {
                                                        continue;
                                                    }
                                                }
                                                else
                                                {
                                                    continue;
                                                }
                                                if (nuevo.GRADO.Contains("°"))
                                                {
                                                    nuevo.GRADO = nuevo.GRADO.Replace("°", string.Empty);
                                                }
                                                if (nuevo.GRADO.Length <= 2 && int.TryParse(nuevo.GRADO, out stx))
                                                {
                                                    student.GRADO = nuevo.GRADO;
                                                }
                                                else
                                                {
                                                    continue;
                                                }

                                                if (nuevo.ASISTIO == "1")
                                                {
                                                    student.ASISTIO = true;
                                                }
                                                else if (nuevo.ASISTIO == "0")
                                                {
                                                    student.ASISTIO = false;
                                                }
                                                else
                                                {
                                                    continue;
                                                }


                                                student.CEDULA = nuevo.CEDULA;
                                                student.COLEGIO_ID = colegio.COLEGIO_ID;
                                                student.PRIMER_NOMBRE = tribunal.Persona.PRIMER_NOMBRE;
                                                student.SEGUNDO_NOMBRE = tribunal.Persona.SEGUNDO_NOMBRE;
                                                student.APELLIDO_PATERNO = tribunal.Persona.APELLIDO_PATERNO;
                                                student.APELLIDO_MATERNO = tribunal.Persona.APELLIDO_MATERNO;
                                                student.PASAPORTE = string.Empty;
                                                student.CEDULA_ACUDIENTE = nuevo.CEDULA_ACUDIENTE;
                                                student.NOMBRE_ACUDIENTE = nuevo.NOMBRE_COMPLETO_ACUDIENTE;
                                                student.PASAPORTE_ACUDIENTE = nuevo.PASAPORTE_ACUDIENTE;
                                                student.FECHA_ULTIMA_MODIFICACION = DateTime.Now;

                                                student.PERIODO = DateTime.Now.Year;



                                                if (!String.IsNullOrWhiteSpace(nuevo.GRUPO))
                                                {
                                                    if (nuevo.GRUPO.Length <= 4)
                                                    {
                                                        student.GRUPO = nuevo.GRUPO;
                                                    }
                                                }
                                                if (nuevo.BACHILLERATO != null)
                                                
                                                {
                                                    if (nuevo.BACHILLERATO.Length <= 50)
                                                    {
                                                        student.BACHILLERATO = nuevo.BACHILLERATO;
                                                    }
                                                }

                                                Db.Estudiantes.Add(student);
                                                Db.CargaEstudiantesAsistencias.Remove(nuevo);
                                                await Db.SaveChangesAsync();
                                                count++;
                                            }
                                            else
                                            {

                                                var resultDate = new DateTime();
                                                if (DateTime.TryParse(tribunal.Persona.FECHA_NACIMIENTO, out resultDate))
                                                {
                                                    int edad = DateTime.Today.Year - resultDate.Date.Year;

                                                    if (resultDate.Date > DateTime.Today.AddYears(-edad)) edad--;

                                                    if (edad <= 21 && edad >= 4)
                                                    {
                                                        valido.FECHA_NACIMIENTO = resultDate.Month + "-" + resultDate.Day + "-" + resultDate.Year;
                                                    }
                                                    else
                                                    {
                                                        continue;
                                                    }
                                                }
                                                else
                                                {
                                                    continue;
                                                }
                                                if (nuevo.GRADO.Contains("°"))
                                                {
                                                    nuevo.GRADO = nuevo.GRADO.Replace("°", string.Empty);
                                                }
                                                if (nuevo.GRADO.Length <= 2 && int.TryParse(nuevo.GRADO, out stx))
                                                {
                                                    valido.GRADO = nuevo.GRADO;
                                                }
                                                else
                                                {
                                                    continue;
                                                }

                                                if (nuevo.ASISTIO == "1")
                                                {
                                                    valido.ASISTIO = true;
                                                }
                                                else if (nuevo.ASISTIO == "0")
                                                {
                                                    valido.ASISTIO = false;
                                                }
                                                else
                                                {
                                                    continue;
                                                }

                                                valido.COLEGIO_ID = colegio.COLEGIO_ID;

                                                valido.PRIMER_NOMBRE = tribunal.Persona.PRIMER_NOMBRE;
                                                valido.SEGUNDO_NOMBRE = tribunal.Persona.SEGUNDO_NOMBRE;
                                                valido.APELLIDO_PATERNO = tribunal.Persona.APELLIDO_PATERNO;
                                                valido.APELLIDO_MATERNO = tribunal.Persona.APELLIDO_MATERNO;

                                                valido.CEDULA_ACUDIENTE = nuevo.CEDULA_ACUDIENTE;
                                                valido.NOMBRE_ACUDIENTE = nuevo.NOMBRE_COMPLETO_ACUDIENTE;
                                                valido.PASAPORTE_ACUDIENTE = nuevo.PASAPORTE_ACUDIENTE;
                                                valido.FECHA_ULTIMA_MODIFICACION = DateTime.Now;
                                                valido.PERIODO = DateTime.Now.Year;

                                                if (!String.IsNullOrWhiteSpace(nuevo.GRUPO))
                                                {
                                                    if (nuevo.GRUPO.Length <= 4)
                                                    {
                                                        valido.GRUPO = nuevo.GRUPO;
                                                    }
                                                }
                                                if (nuevo.BACHILLERATO != null)

                                                {
                                                    if (nuevo.BACHILLERATO.Length <= 50)
                                                    {
                                                        valido.BACHILLERATO = nuevo.BACHILLERATO;
                                                    }
                                                }


                                                Db.CargaEstudiantesAsistencias.Remove(nuevo);
                                                await Db.SaveChangesAsync();
                                                count++;
                                            }
                                        }

                                    }
                                    else
                                    {
                                        Db.CargaEstudiantesAsistencias.Remove(nuevo);
                                        await Db.SaveChangesAsync();
                                    }
                                }
                                else if (!String.IsNullOrWhiteSpace(nuevo.PASAPORTE))
                                {
                                    var estudiante = Db.Estudiantes.Where(x => x.PASAPORTE == nuevo.PASAPORTE).FirstOrDefault();

                                    if (estudiante == null)
                                    {
                                        var student = new Estudiante();
                                        var resultDate = new DateTime();

                                        if (DateTime.TryParse(nuevo.FECHA_NACIMIENTO, out resultDate))
                                        {
                                            int edad = DateTime.Today.Year - resultDate.Date.Year;

                                            if (resultDate.Date > DateTime.Today.AddYears(-edad)) edad--;

                                            if (edad <= 21 && edad >= 4)
                                            {
                                                student.FECHA_NACIMIENTO = resultDate.Month + "-" + resultDate.Day + "-" + resultDate.Year;
                                            }
                                            else
                                            {
                                                continue;
                                            }
                                        }
                                        else
                                        {
                                            continue;
                                        }

                                        if (nuevo.GRADO.Contains("°"))
                                        {
                                            nuevo.GRADO = nuevo.GRADO.Replace("°", string.Empty);
                                        }
                                        if (nuevo.GRADO.Length <= 2 && int.TryParse(nuevo.GRADO, out stx))
                                        {
                                            student.GRADO = nuevo.GRADO;
                                        }
                                        else
                                        {
                                            continue;
                                        }

                                        if (nuevo.ASISTIO == "1")
                                        {
                                            student.ASISTIO = true;
                                        }
                                        else if (nuevo.ASISTIO == "0")
                                        {
                                            student.ASISTIO = false;
                                        }
                                        else
                                        {
                                            continue;
                                        }


                                        student.CEDULA = null;
                                        student.COLEGIO_ID = colegio.COLEGIO_ID;
                                        student.PRIMER_NOMBRE = nuevo.PRIMER_NOMBRE;
                                        student.SEGUNDO_NOMBRE = nuevo.SEGUNDO_NOMBRE;
                                        student.APELLIDO_PATERNO = nuevo.APELLIDO_PATERNO;
                                        student.APELLIDO_MATERNO = nuevo.APELLIDO_MATERNO;

                                        student.PASAPORTE = nuevo.PASAPORTE;
                                        student.CEDULA_ACUDIENTE = nuevo.CEDULA_ACUDIENTE;
                                        student.NOMBRE_ACUDIENTE = nuevo.NOMBRE_COMPLETO_ACUDIENTE;
                                        student.PASAPORTE_ACUDIENTE = nuevo.PASAPORTE_ACUDIENTE;
                                        student.FECHA_ULTIMA_MODIFICACION = DateTime.Now;
                                        student.PERIODO = DateTime.Now.Year;




                                        if (!String.IsNullOrWhiteSpace(nuevo.GRUPO))
                                        {
                                            if (nuevo.GRUPO.Length <= 4)
                                            {
                                                student.GRUPO = nuevo.GRUPO;
                                            }
                                        }
                                        if (nuevo.BACHILLERATO != null)
                                       
                                        {
                                            if (nuevo.BACHILLERATO.Length <= 50)
                                            {
                                                student.BACHILLERATO = nuevo.BACHILLERATO;
                                            }
                                        }

                                        Db.Estudiantes.Add(student);
                                        Db.CargaEstudiantesAsistencias.Remove(nuevo);
                                        await Db.SaveChangesAsync();
                                        count++;
                                    }
                                    else
                                    {
                                        var resultDate = new DateTime();
                                        if (DateTime.TryParse(nuevo.FECHA_NACIMIENTO, out resultDate))
                                        {
                                            int edad = DateTime.Today.Year - resultDate.Date.Year;

                                            if (resultDate.Date > DateTime.Today.AddYears(-edad)) edad--;

                                            if (edad <= 21 && edad >= 4)
                                            {
                                                estudiante.FECHA_NACIMIENTO = resultDate.Month + "-" + resultDate.Day + "-" + resultDate.Year;
                                            }
                                            else
                                            {
                                                continue;
                                            }
                                        }
                                        else
                                        {
                                            continue;
                                        }
                                        if (nuevo.GRADO.Contains("°"))
                                        {
                                            nuevo.GRADO = nuevo.GRADO.Replace("°", string.Empty);
                                        }
                                        if (nuevo.GRADO.Length <= 2 && int.TryParse(nuevo.GRADO, out stx))
                                        {
                                            estudiante.GRADO = nuevo.GRADO;
                                        }
                                        else
                                        {
                                            continue;
                                        }

                                        if (nuevo.ASISTIO == "1")
                                        {
                                            estudiante.ASISTIO = true;
                                        }
                                        else if (nuevo.ASISTIO == "0")
                                        {
                                            estudiante.ASISTIO = false;
                                        }
                                        else
                                        {
                                            continue;
                                        }

                                        estudiante.COLEGIO_ID = colegio.COLEGIO_ID;

                                        estudiante.PRIMER_NOMBRE = nuevo.PRIMER_NOMBRE;
                                        estudiante.SEGUNDO_NOMBRE = nuevo.SEGUNDO_NOMBRE;
                                        estudiante.APELLIDO_PATERNO = nuevo.APELLIDO_PATERNO;
                                        estudiante.APELLIDO_MATERNO = nuevo.APELLIDO_MATERNO;

                                        estudiante.CEDULA_ACUDIENTE = nuevo.CEDULA_ACUDIENTE;
                                        estudiante.NOMBRE_ACUDIENTE = nuevo.NOMBRE_COMPLETO_ACUDIENTE;
                                        estudiante.PASAPORTE_ACUDIENTE = nuevo.PASAPORTE_ACUDIENTE;
                                        estudiante.FECHA_ULTIMA_MODIFICACION = DateTime.Now;
                                        estudiante.PERIODO = DateTime.Now.Year;



                                        if (!String.IsNullOrWhiteSpace(nuevo.GRUPO))
                                        {
                                            if (nuevo.GRUPO.Length <= 4)
                                            {
                                                estudiante.GRUPO = nuevo.GRUPO;
                                            }
                                        }
                                        if (nuevo.BACHILLERATO != null)
                                       
                                        {
                                            if (nuevo.BACHILLERATO.Length <= 50)
                                            {
                                                estudiante.BACHILLERATO = nuevo.BACHILLERATO;
                                            }
                                        }


                                        Db.CargaEstudiantesAsistencias.Remove(nuevo);
                                        await Db.SaveChangesAsync();
                                        count++;
                                    }
                                }
                            }
                        }
                        else if (cargapago == 3)
                        {
                            conString = ConfigurationManager.ConnectionStrings["DBCS"].ConnectionString;
                            using (SqlConnection con = new SqlConnection(conString))
                            {
                                using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                                {

                                    sqlBulkCopy.DestinationTableName = "CargaEstudiantesNota";
                                    sqlBulkCopy.ColumnMappings.Add(1, "CEDULA");
                                    sqlBulkCopy.ColumnMappings.Add(2, "PRIMER_NOMBRE");
                                    sqlBulkCopy.ColumnMappings.Add(3, "SEGUNDO_NOMBRE");
                                    sqlBulkCopy.ColumnMappings.Add(4, "APELLIDO_PATERNO");
                                    sqlBulkCopy.ColumnMappings.Add(5, "APELLIDO_MATERNO");
                                    sqlBulkCopy.ColumnMappings.Add(6, "PASAPORTE");
                                    sqlBulkCopy.ColumnMappings.Add(7, "GRADO");
                                    sqlBulkCopy.ColumnMappings.Add(8, "GRUPO");
                                    sqlBulkCopy.ColumnMappings.Add(9, "BACHILLERATO");
                                    sqlBulkCopy.ColumnMappings.Add(10, "FECHA_NACIMIENTO");
                                    sqlBulkCopy.ColumnMappings.Add(11, "CEDULA_ACUDIENTE");
                                    sqlBulkCopy.ColumnMappings.Add(12, "NOMBRE_COMPLETO_ACUDIENTE");
                                    sqlBulkCopy.ColumnMappings.Add(13, "PASAPORTE_ACUDIENTE");
                                    sqlBulkCopy.ColumnMappings.Add(14, "PROMEDIO");
                                    sqlBulkCopy.ColumnMappings.Add(15, "MATERIA_MAS_BAJA");
                                    sqlBulkCopy.ColumnMappings.Add(16, "NOTA_MAS_BAJA");
                                    sqlBulkCopy.ColumnMappings.Add(17, "FECHA");
                                    sqlBulkCopy.ColumnMappings.Add(18, "USUARIO");
                                    con.Open();
                                    sqlBulkCopy.WriteToServer(dt);
                                    con.Close();
                                }
                            }

                            var carga = Db.CargaEstudiantesNotas.Where(x => x.USUARIO == user).ToList();


                            var rx = new Regex(@"^(PE|E|N|[23456789](?:AV|PI)?|1[0123]?(?:AV|PI)?)-(?:[1-9]|[1-9][0-9]{1,3})-(?:[1-9]|[1-9][0-9]{1,5})$", RegexOptions.IgnoreCase);

                            foreach (var nuevo in carga)
                            {
                                if (!String.IsNullOrWhiteSpace(nuevo.CEDULA_ACUDIENTE))
                                {
                                    nuevo.CEDULA_ACUDIENTE = nuevo.CEDULA_ACUDIENTE.ToUpper().Trim();

                                    if (!rx.IsMatch(nuevo.CEDULA_ACUDIENTE))
                                    {
                                        continue;
                                    }
                                }
                                else if (String.IsNullOrWhiteSpace(nuevo.PASAPORTE_ACUDIENTE))
                                {
                                    continue;
                                }
                                if (String.IsNullOrWhiteSpace(nuevo.FECHA_NACIMIENTO) || String.IsNullOrWhiteSpace(nuevo.GRADO) || String.IsNullOrWhiteSpace(nuevo.NOMBRE_COMPLETO_ACUDIENTE))
                                {
                                    continue;
                                }
                                nuevo.GRADO = nuevo.GRADO.ToUpper().Trim();

                                if (!String.IsNullOrWhiteSpace(nuevo.CEDULA))
                                {
                                    nuevo.CEDULA = nuevo.CEDULA.ToUpper().Trim();

                                    if (rx.IsMatch(nuevo.CEDULA))
                                    {
                                        var valido = Db.Estudiantes.Where(x => x.CEDULA == nuevo.CEDULA).FirstOrDefault();
                                        var tribunal = await ValidarEstudiantes(nuevo.CEDULA);
                                        if (tribunal.MensajeError == null)
                                        {
                                            if (valido == null)
                                            {
                                                var student = new Estudiante();

                                                var resultDate = new DateTime();
                                                if (DateTime.TryParse(tribunal.Persona.FECHA_NACIMIENTO, out resultDate))
                                                {
                                                    int edad = DateTime.Today.Year - resultDate.Date.Year;

                                                    if (resultDate.Date > DateTime.Today.AddYears(-edad)) edad--;

                                                    if (edad <= 21 && edad >= 4)
                                                    {
                                                        student.FECHA_NACIMIENTO = resultDate.Month + "-" + resultDate.Day + "-" + resultDate.Year;
                                                    }
                                                    else
                                                    {
                                                        continue;
                                                    }
                                                }
                                                else
                                                {
                                                    continue;
                                                }
                                                if (nuevo.GRADO.Contains("°"))
                                                {
                                                    nuevo.GRADO = nuevo.GRADO.Replace("°", string.Empty);
                                                }
                                                if (nuevo.GRADO.Length <= 2 && int.TryParse(nuevo.GRADO, out stx))
                                                {
                                                    student.GRADO = nuevo.GRADO;
                                                }
                                                else
                                                {
                                                    continue;
                                                }

                                                var resultprom = new Decimal();
                                                var resultnota = new Decimal();
                                                if (nuevo.PROMEDIO != null && nuevo.NOTA_MAS_BAJA != null && nuevo.MATERIA_MAS_BAJA != null)
                                                {
                                                    if (Decimal.TryParse(nuevo.PROMEDIO, out resultprom) && Decimal.TryParse(nuevo.NOTA_MAS_BAJA, out resultnota))
                                                    {
                                                        if ((resultprom <= 5.0M && resultprom >= 1.0M && resultnota <= 5.0M && resultnota >= 1.0M))
                                                        {
                                                            if (resultnota > resultprom)
                                                            {
                                                                continue;
                                                            }
                                                            else
                                                            {
                                                                student.PROMEDIO = resultprom;
                                                                student.MATERIA_MAS_BAJA = nuevo.MATERIA_MAS_BAJA;
                                                                student.NOTA_MAS_BAJA = resultnota;
                                                            }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        continue;
                                                    }

                                                }
                                                else
                                                {
                                                    continue;
                                                }



                                                student.CEDULA = nuevo.CEDULA;
                                                student.COLEGIO_ID = colegio.COLEGIO_ID;
                                                student.PRIMER_NOMBRE = tribunal.Persona.PRIMER_NOMBRE;
                                                student.SEGUNDO_NOMBRE = tribunal.Persona.SEGUNDO_NOMBRE;
                                                student.APELLIDO_PATERNO = tribunal.Persona.APELLIDO_PATERNO;
                                                student.APELLIDO_MATERNO = tribunal.Persona.APELLIDO_MATERNO;
                                                student.PASAPORTE = string.Empty;
                                                student.CEDULA_ACUDIENTE = nuevo.CEDULA_ACUDIENTE;
                                                student.NOMBRE_ACUDIENTE = nuevo.NOMBRE_COMPLETO_ACUDIENTE;
                                                student.PASAPORTE_ACUDIENTE = nuevo.PASAPORTE_ACUDIENTE;
                                                student.FECHA_ULTIMA_MODIFICACION = DateTime.Now;
                                                student.PERIODO = DateTime.Now.Year;



                                                if (!String.IsNullOrWhiteSpace(nuevo.GRUPO))
                                                {
                                                    if (nuevo.GRUPO.Length <= 4)
                                                    {
                                                        student.GRUPO = nuevo.GRUPO;
                                                    }
                                                }
                                                if (!String.IsNullOrWhiteSpace(nuevo.BACHILLERATO))
                                                {
                                                    if (nuevo.BACHILLERATO.Length <= 50)
                                                    {
                                                        student.BACHILLERATO = nuevo.BACHILLERATO;
                                                    }
                                                }

                                                Db.Estudiantes.Add(student);
                                                Db.CargaEstudiantesNotas.Remove(nuevo);
                                                await Db.SaveChangesAsync();
                                                count++;
                                            }
                                            else
                                            {

                                                var resultDate = new DateTime();
                                                if (DateTime.TryParse(tribunal.Persona.FECHA_NACIMIENTO, out resultDate))
                                                {
                                                    int edad = DateTime.Today.Year - resultDate.Date.Year;

                                                    if (resultDate.Date > DateTime.Today.AddYears(-edad)) edad--;

                                                    if (edad <= 21 && edad >= 4)
                                                    {
                                                        valido.FECHA_NACIMIENTO = resultDate.Month + "-" + resultDate.Day + "-" + resultDate.Year;
                                                    }
                                                    else
                                                    {
                                                        continue;
                                                    }
                                                }
                                                else
                                                {
                                                    continue;
                                                }
                                                if (nuevo.GRADO.Contains("°"))
                                                {
                                                    nuevo.GRADO = nuevo.GRADO.Replace("°", string.Empty);
                                                }
                                                if (nuevo.GRADO.Length <= 2 && int.TryParse(nuevo.GRADO, out stx))
                                                {
                                                    valido.GRADO = nuevo.GRADO;
                                                }
                                                else
                                                {
                                                    continue;
                                                }

                                                var resultprom = new Decimal();
                                                var resultnota = new Decimal();
                                                if (nuevo.PROMEDIO != null && nuevo.NOTA_MAS_BAJA != null && nuevo.MATERIA_MAS_BAJA != null)
                                                {
                                                    if (Decimal.TryParse(nuevo.PROMEDIO, out resultprom) && Decimal.TryParse(nuevo.NOTA_MAS_BAJA, out resultnota))
                                                    {
                                                        if ((resultprom <= 5.0M && resultprom >= 1.0M && resultnota <= 5.0M && resultnota >= 1.0M))
                                                        {
                                                            if (resultnota > resultprom)
                                                            {
                                                                continue;
                                                            }
                                                            else
                                                            {
                                                                valido.PROMEDIO = resultprom;
                                                                valido.MATERIA_MAS_BAJA = nuevo.MATERIA_MAS_BAJA;
                                                                valido.NOTA_MAS_BAJA = resultnota;
                                                            }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        continue;
                                                    }

                                                }
                                                else
                                                {
                                                    continue;
                                                }


                                                valido.COLEGIO_ID = colegio.COLEGIO_ID;

                                                valido.PRIMER_NOMBRE = tribunal.Persona.PRIMER_NOMBRE;
                                                valido.SEGUNDO_NOMBRE = tribunal.Persona.SEGUNDO_NOMBRE;
                                                valido.APELLIDO_PATERNO = tribunal.Persona.APELLIDO_PATERNO;
                                                valido.APELLIDO_MATERNO = tribunal.Persona.APELLIDO_MATERNO;

                                                valido.CEDULA_ACUDIENTE = nuevo.CEDULA_ACUDIENTE;
                                                valido.NOMBRE_ACUDIENTE = nuevo.NOMBRE_COMPLETO_ACUDIENTE;
                                                valido.PASAPORTE_ACUDIENTE = nuevo.PASAPORTE_ACUDIENTE;
                                                valido.FECHA_ULTIMA_MODIFICACION = DateTime.Now;
                                                valido.PERIODO = DateTime.Now.Year;

                                                if (!String.IsNullOrWhiteSpace(nuevo.GRUPO))
                                                {
                                                    if (nuevo.GRUPO.Length <= 4)
                                                    {
                                                        valido.GRUPO = nuevo.GRUPO;
                                                    }
                                                }
                                                if (!String.IsNullOrWhiteSpace(nuevo.BACHILLERATO))
                                                {
                                                    if (nuevo.BACHILLERATO.Length <= 50)
                                                    {
                                                        valido.BACHILLERATO = nuevo.BACHILLERATO;
                                                    }
                                                }


                                                Db.CargaEstudiantesNotas.Remove(nuevo);
                                                await Db.SaveChangesAsync();
                                                count++;
                                            }
                                        }

                                    }
                                    else
                                    {
                                        Db.CargaEstudiantesNotas.Remove(nuevo);
                                        await Db.SaveChangesAsync();
                                    }
                                }
                                else if (!String.IsNullOrWhiteSpace(nuevo.PASAPORTE))
                                {
                                    var estudiante = Db.Estudiantes.Where(x => x.PASAPORTE == nuevo.PASAPORTE).FirstOrDefault();

                                    if (estudiante == null)
                                    {
                                        var student = new Estudiante();
                                        var resultDate = new DateTime();

                                        if (DateTime.TryParse(nuevo.FECHA_NACIMIENTO, out resultDate))
                                        {
                                            int edad = DateTime.Today.Year - resultDate.Date.Year;

                                            if (resultDate.Date > DateTime.Today.AddYears(-edad)) edad--;

                                            if (edad <= 21 && edad >= 4)
                                            {
                                                student.FECHA_NACIMIENTO = resultDate.Month + "-" + resultDate.Day + "-" + resultDate.Year;
                                            }
                                            else
                                            {
                                                continue;
                                            }
                                        }
                                        else
                                        {
                                            continue;
                                        }

                                        if (nuevo.GRADO.Contains("°"))
                                        {
                                            nuevo.GRADO = nuevo.GRADO.Replace("°", string.Empty);
                                        }
                                        if (nuevo.GRADO.Length <= 2 && int.TryParse(nuevo.GRADO, out stx))
                                        {
                                            student.GRADO = nuevo.GRADO;
                                        }
                                        else
                                        {
                                            continue;
                                        }

                                        var resultprom = new Decimal();
                                        var resultnota = new Decimal();
                                        if (nuevo.PROMEDIO != null && nuevo.NOTA_MAS_BAJA != null && nuevo.MATERIA_MAS_BAJA != null)
                                        {
                                            if (Decimal.TryParse(nuevo.PROMEDIO, out resultprom) && Decimal.TryParse(nuevo.NOTA_MAS_BAJA, out resultnota))
                                            {
                                                if ((resultprom <= 5.0M && resultprom >= 1.0M && resultnota <= 5.0M && resultnota >= 1.0M))
                                                {
                                                    if (resultnota > resultprom)
                                                    {
                                                        continue;
                                                    }
                                                    else
                                                    {
                                                        student.PROMEDIO = resultprom;
                                                        student.MATERIA_MAS_BAJA = nuevo.MATERIA_MAS_BAJA;
                                                        student.NOTA_MAS_BAJA = resultnota;
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                continue;
                                            }

                                        }
                                        else
                                        {
                                            continue;
                                        }


                                        student.CEDULA = null;
                                        student.COLEGIO_ID = colegio.COLEGIO_ID;
                                        student.PRIMER_NOMBRE = nuevo.PRIMER_NOMBRE;
                                        student.SEGUNDO_NOMBRE = nuevo.SEGUNDO_NOMBRE;
                                        student.APELLIDO_PATERNO = nuevo.APELLIDO_PATERNO;
                                        student.APELLIDO_MATERNO = nuevo.APELLIDO_MATERNO;

                                        student.PASAPORTE = nuevo.PASAPORTE;
                                        student.CEDULA_ACUDIENTE = nuevo.CEDULA_ACUDIENTE;
                                        student.NOMBRE_ACUDIENTE = nuevo.NOMBRE_COMPLETO_ACUDIENTE;
                                        student.PASAPORTE_ACUDIENTE = nuevo.PASAPORTE_ACUDIENTE;
                                        student.FECHA_ULTIMA_MODIFICACION = DateTime.Now;
                                        student.PERIODO = DateTime.Now.Year;




                                        if (!String.IsNullOrWhiteSpace(nuevo.GRUPO))
                                        {
                                            if (nuevo.GRUPO.Length <= 4)
                                            {
                                                student.GRUPO = nuevo.GRUPO;
                                            }
                                        }
                                        if (!String.IsNullOrWhiteSpace(nuevo.BACHILLERATO))
                                        {
                                            if (nuevo.BACHILLERATO.Length <= 50)
                                            {
                                                student.BACHILLERATO = nuevo.BACHILLERATO;
                                            }
                                        }

                                        Db.Estudiantes.Add(student);
                                        Db.CargaEstudiantesNotas.Remove(nuevo);
                                        await Db.SaveChangesAsync();
                                        count++;
                                    }
                                    else
                                    {
                                        var resultDate = new DateTime();
                                        if (DateTime.TryParse(nuevo.FECHA_NACIMIENTO, out resultDate))
                                        {
                                            int edad = DateTime.Today.Year - resultDate.Date.Year;

                                            if (resultDate.Date > DateTime.Today.AddYears(-edad)) edad--;

                                            if (edad <= 21 && edad >= 4)
                                            {
                                                estudiante.FECHA_NACIMIENTO = resultDate.Month + "-" + resultDate.Day + "-" + resultDate.Year;
                                            }
                                            else
                                            {
                                                continue;
                                            }
                                        }
                                        else
                                        {
                                            continue;
                                        }
                                        if (nuevo.GRADO.Contains("°"))
                                        {
                                            nuevo.GRADO = nuevo.GRADO.Replace("°", string.Empty);
                                        }
                                        if (nuevo.GRADO.Length <= 2 && int.TryParse(nuevo.GRADO, out stx))
                                        {
                                            estudiante.GRADO = nuevo.GRADO;
                                        }
                                        else
                                        {
                                            continue;
                                        }

                                        var resultprom = new Decimal();
                                        var resultnota = new Decimal();
                                        if (nuevo.PROMEDIO != null && nuevo.NOTA_MAS_BAJA != null && nuevo.MATERIA_MAS_BAJA != null)
                                        {
                                            if (Decimal.TryParse(nuevo.PROMEDIO, out resultprom) && Decimal.TryParse(nuevo.NOTA_MAS_BAJA, out resultnota))
                                            {
                                                if ((resultprom <= 5.0M && resultprom >= 1.0M && resultnota <= 5.0M && resultnota >= 1.0M))
                                                {
                                                    if (resultnota > resultprom)
                                                    {
                                                        continue;
                                                    }
                                                    else
                                                    {
                                                        estudiante.PROMEDIO = resultprom;
                                                        estudiante.MATERIA_MAS_BAJA = nuevo.MATERIA_MAS_BAJA;
                                                        estudiante.NOTA_MAS_BAJA = resultnota;
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                continue;
                                            }

                                        }
                                        else
                                        {
                                            continue;
                                        }

                                        estudiante.COLEGIO_ID = colegio.COLEGIO_ID;

                                        estudiante.PRIMER_NOMBRE = nuevo.PRIMER_NOMBRE;
                                        estudiante.SEGUNDO_NOMBRE = nuevo.SEGUNDO_NOMBRE;
                                        estudiante.APELLIDO_PATERNO = nuevo.APELLIDO_PATERNO;
                                        estudiante.APELLIDO_MATERNO = nuevo.APELLIDO_MATERNO;

                                        estudiante.CEDULA_ACUDIENTE = nuevo.CEDULA_ACUDIENTE;
                                        estudiante.NOMBRE_ACUDIENTE = nuevo.NOMBRE_COMPLETO_ACUDIENTE;
                                        estudiante.PASAPORTE_ACUDIENTE = nuevo.PASAPORTE_ACUDIENTE;
                                        estudiante.FECHA_ULTIMA_MODIFICACION = DateTime.Now;
                                        estudiante.PERIODO = DateTime.Now.Year;



                                        if (!String.IsNullOrWhiteSpace(nuevo.GRUPO))
                                        {
                                            if (nuevo.GRUPO.Length <= 4)
                                            {
                                                estudiante.GRUPO = nuevo.GRUPO;
                                            }
                                        }
                                        if (!String.IsNullOrWhiteSpace(nuevo.BACHILLERATO))
                                        {
                                            if (nuevo.BACHILLERATO.Length <= 50)
                                            {
                                                estudiante.BACHILLERATO = nuevo.BACHILLERATO;
                                            }
                                        }


                                        Db.CargaEstudiantesNotas.Remove(nuevo);
                                        await Db.SaveChangesAsync();
                                        count++;
                                    }
                                }
                            }
                        }


                    }

                    if (count != 0)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }

            catch (DbEntityValidationException e)
            {
                string ex = "";
                foreach (var eve in e.EntityValidationErrors)
                {
                    ex += ("Entity of type \"{0}\" in state \"{1}\" has the following validation errors:",
                         eve.Entry.Entity.GetType().Name, eve.Entry.State);
                    foreach (var ve in eve.ValidationErrors)
                    {
                        ex += ("- Property: \"{0}\", Error: \"{1}\"",
                            ve.PropertyName, ve.ErrorMessage);
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                await ErrorLog(ex.ToString());
                return false;
            }

        }