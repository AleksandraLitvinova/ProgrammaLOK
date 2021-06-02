using System;
using System.Collections.Generic;
using System.Text;

namespace ProgrammaLOK
{
    class EmployeeVaccinationRelation
    {
        int idEmployee;
        int idVaccination;
        int idStatus;
        public int dateVaccination;
        string datePlanVaccination;

        public EmployeeVaccinationRelation(int idEmployee, int idVaccination)
        {
            this.idEmployee = idEmployee;
            this.idVaccination = idVaccination;
        }
    }
}
