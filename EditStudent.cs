        public async Task EditarRegistro(int id, string ceddient, string pasdient, string nomdient, string bachi, string grado, string grupo, bool? asistio, decimal? promedio, string materia, decimal? nota)
        {
            try
            {
                using (SCDPEntities Db = new SCDPEntities())
                {
                    var student = Db.Estudiantes.Where(x => x.ESTUDIANTE_ID == id).FirstOrDefault();


                    if (!String.IsNullOrWhiteSpace(grado) && !String.IsNullOrWhiteSpace(nomdient))
                    {

                        var resultDate = new DateTime();
                        if (DateTime.TryParse(student.FECHA_NACIMIENTO, out resultDate))
                        {
                            int edad = DateTime.Today.Year - resultDate.Date.Year;

                            if (resultDate.Date > DateTime.Today.AddYears(-edad)) edad--;

                            if (edad > 21 || edad < 4)
                            {
                                return;
                            }
                        }
                        else
                        {
                            return;
                        }



                        var rx = new Regex(@"^(PE|E|N|[23456789](?:AV|PI)?|1[0123]?(?:AV|PI)?)-(?:[1-9]|[1-9][0-9]{1,3})-(?:[1-9]|[1-9][0-9]{1,5})$", RegexOptions.IgnoreCase);

                        if (!String.IsNullOrWhiteSpace(ceddient))
                        {
                            ceddient = ceddient.ToUpper();

                            if (!rx.IsMatch(ceddient))
                            {
                                return;
                            }
                        }
                        else if (String.IsNullOrWhiteSpace(pasdient))
                        {
                            return;
                        }

                        if (grado.Contains("°"))
                        {
                            grado = grado.Replace("°", string.Empty);
                        }
                        if (grado.Length <= 2 && int.TryParse(grado, out int stx))
                        {
                            student.GRADO = grado;
                        }
                        else
                        {
                            return;
                        }

                        student.CEDULA_ACUDIENTE = ceddient;
                        student.PASAPORTE_ACUDIENTE = pasdient;
                        student.NOMBRE_ACUDIENTE = nomdient;
                        student.BACHILLERATO = bachi;

                        student.GRUPO = grupo;
                        student.ASISTIO = asistio;
                        student.MATERIA_MAS_BAJA = materia;
                        student.PERIODO = DateTime.Now.Year;
                        student.FECHA_ULTIMA_MODIFICACION = DateTime.Now;
                        if(promedio != null && nota != null)
                        {
                            if ((promedio <= 5.0M && promedio >= 1.0M && nota <= 5.0M && nota >= 1.0M))
                            {
                                if (nota > promedio)
                                {
                                    return;
                                }
                                else
                                {
                                    student.PROMEDIO = promedio;

                                    student.NOTA_MAS_BAJA = nota;
                                }
                            }
                        }
                        else
                        {
                            student.PROMEDIO = null;

                            student.NOTA_MAS_BAJA = null;
                        }
                      
                        await Db.SaveChangesAsync();
                    }


                }
            }
            catch (Exception ex)
            {
                await ErrorLog(ex.ToString());
            }
        }