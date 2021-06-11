using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;

namespace ProgrammaLOK
{
    class StatusVaccination
    {
        public int id;
        public string name;
        List<string> status_variants = new List<string>(); 

        public StatusVaccination(int idStatus, string nameStatus, List<string> variants =null)
        {
            this.id = idStatus;
            this.name = nameStatus;
            if(variants==null)
            {
                status_variants.Add(nameStatus);
            }
            else
            {
                this.status_variants = variants;
            }
            
        }
        public bool Equals(string status)
        {
            return status_variants.Select(n=>n.ToLower()).Contains(status.ToLower());
        }
    }
}
