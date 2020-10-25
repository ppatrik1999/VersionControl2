using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Mikroszimulácuó.Entities;



namespace Mikroszimulácuó.Entities
{
    public enum Gender
    {
        Male = 1,
        Female = 2,
    }

    public class Person
    {
        public int BirthYear { get; set; }
        public Gender Gender { get; set; }
        public int NbrOfChildren { get; set; }
        public bool IsAlive { get; set; }
       


        public Person()
        {
            IsAlive = true;
        }
    }



}
