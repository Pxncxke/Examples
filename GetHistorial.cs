        public List<CentroEducativo> GetHistorial(string centroid, int option, int select)
        {
            try
            {
                using ACEContext context = new();
                List<CentroEducativo> centro = select switch
                {
                    1 => option switch
                    {
                        1 => context.CentroEducativos.Include(r => r.ServicioInternets).ThenInclude(t => t.Plan).Include(r => r.Matriculas).Include(r => r.ServicioInternets).ThenInclude(t => t.Plan).Include(r => r.ProyectoAsignados).ThenInclude(t => t.Proyecto).Include(r => r.ElectricidadAsignada).ThenInclude(t => t.Electricidad).Include(t => t.ElectricidadAsignada).Where(x => x.CentroId == centroid).ToList(),
                        2 => context.CentroEducativos.Include(r => r.ServicioInternets).ThenInclude(t => t.Plan).Where(x => x.CentroId == centroid).ToList(),
                        3 => context.CentroEducativos.Include(r => r.ProyectoAsignados).ThenInclude(t => t.Proyecto).Where(x => x.CentroId == centroid).ToList(),
                        4 => context.CentroEducativos.Include(r => r.ElectricidadAsignada).ThenInclude(t => t.Electricidad).Include(t => t.ElectricidadAsignada).Where(x => x.CentroId == centroid).ToList(),
                        5 => context.CentroEducativos.Include(r => r.Matriculas).Where(x => x.CentroId == centroid).ToList(),
                        _ => new List<CentroEducativo>(),
                    },
                    2 => option switch
                    {
                        1 => context.CentroEducativos.Include(r => r.ServicioInternets).ThenInclude(t => t.Plan).Include(r => r.Matriculas).Include(r => r.ServicioInternets).ThenInclude(t => t.Plan).Include(r => r.ProyectoAsignados).ThenInclude(t => t.Proyecto).Include(r => r.ElectricidadAsignada).ThenInclude(t => t.Electricidad).Include(t => t.ElectricidadAsignada).Where(x => x.NombreCentroEducativo.Contains(centroid)).ToList(),
                        2 => context.CentroEducativos.Include(r => r.ServicioInternets).ThenInclude(t => t.Plan).Where(x => x.NombreCentroEducativo.Contains(centroid)).ToList(),
                        3 => context.CentroEducativos.Include(r => r.ProyectoAsignados).ThenInclude(t => t.Proyecto).Where(x => x.NombreCentroEducativo.Contains(centroid)).ToList(),
                        4 => context.CentroEducativos.Include(r => r.ElectricidadAsignada).ThenInclude(t => t.Electricidad).Include(t => t.ElectricidadAsignada).Where(x => x.NombreCentroEducativo.Contains(centroid)).ToList(),
                        5 => context.CentroEducativos.Include(r => r.Matriculas).Where(x => x.NombreCentroEducativo.Contains(centroid)).ToList(),
                        _ => new List<CentroEducativo>(),
                    },
                    _ => new List<CentroEducativo>(),
                };

                return centro;
            }
            catch (Exception ex)
            {
                ErrorLogTxt(ex.ToString());
                return new List<CentroEducativo>();
            }
        }