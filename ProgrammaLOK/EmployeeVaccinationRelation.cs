using System;
using System.Collections.Generic;
using System.Text;

namespace ProgrammaLOK
{
    class EmployeeVaccinationRelation
    {
        int idEmployee;
        int idVaccination;
        public int idStatus;
        public int dateVaccination;
        string datePlanVaccination;
        public DateTime dt;

        public EmployeeVaccinationRelation(int idEmployee, int idVaccination)
        {
            this.idEmployee = idEmployee;
            this.idVaccination = idVaccination;
        }
    }
}
